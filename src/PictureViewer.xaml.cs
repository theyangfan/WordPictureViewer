using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using F = System.Windows.Forms;
using C = System.Windows.Controls;
using System.Drawing;
using System.Drawing.Imaging;
using Word = Microsoft.Office.Interop.Word;
using System.Threading;

namespace WordPictureViewer
{
    /// <summary>
    /// UserControl1.xaml 的交互逻辑
    /// </summary>
    public partial class PictureViewer : Window
    {
        #region Private Members
        private const int MINIMUM_SCALE = 50; // 50%
        private Bitmap _bitmap = null;
        private double _oriWidth;
        private double _oriHeight;
        private int _scale = 100; // 100%
        private int _scaleStep = 10;
        private bool _isZoomIn = true;

        private System.Windows.Point _pressedPoint;
        private bool _isPressed = false;
        private double _lastXOffset = 0;
        private double _lastYOffset = 0;
        #endregion

        #region Constructor
        public PictureViewer()
        {
            InitializeComponent();

            ResizeMode = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            WindowState = WindowState.Maximized;
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;

            InitUI();

            // Mouse events
            MouseWheel += PictureViewer_MouseWheel;

            uiContent.MouseDown += PictureViewer_MouseDown;

            uiContent.MouseUp += PictureViewer_MouseUp;

            uiContent.MouseMove += PictureViewer_MouseMove;

            _pressedPoint = new System.Windows.Point(0, 0);
        }


        #endregion

        #region Public Method
        /// <summary>
        /// Show inline shape.
        /// </summary>
        /// <param name="sel"></param>
        public void ShowInlineShapeSelection(Word.Selection sel)
        {
            try
            {
                var shape = sel.InlineShapes[1];
                // get the bytes of the shape
                var bits = (byte[])sel.EnhMetaFileBits;
                using (MemoryStream stream = new MemoryStream(bits))
                {
                    Bitmap srcBmp = new Bitmap(stream);
                    // get the available width
                    _oriWidth = srcBmp.Height * (shape.Width / shape.Height);
                    _oriHeight = srcBmp.Height;
                    // crop the available area
                    _bitmap = Crop(srcBmp, 0, 0, (int)_oriWidth, (int)_oriHeight);
                    lblPicSize.Content = $"{_bitmap.Width} x {_bitmap.Height}";
                    double ratio = (double)_oriWidth / _oriHeight;
                    double screenW = SystemParameters.PrimaryScreenWidth;
                    double screenH = SystemParameters.PrimaryScreenHeight;
                    if (_oriWidth > screenW)
                    {
                        _oriWidth = screenW;
                        _oriHeight = _oriWidth / ratio;
                    }
                    if (_oriHeight > screenH)
                    {
                        _oriHeight = screenH;
                        _oriWidth = _oriHeight * ratio;
                    }
                    _oriWidth *= 0.8;
                    _oriHeight *= 0.8;
                    uiImage.Width = _oriWidth;
                    uiImage.Height = _oriHeight;

                    // set image source
                    using (MemoryStream ms = new MemoryStream())
                    {
                        _bitmap.Save(ms, ImageFormat.Png);
                        var bytes = ms.GetBuffer();
                        BitmapImage bi = new BitmapImage();
                        bi.BeginInit();
                        bi.StreamSource = new MemoryStream(bytes);
                        bi.EndInit();
                        uiImage.Source = bi;
                    }
                    // loading animation
                    Scale(0, 1, _oriWidth / 2, _oriHeight / 2);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show($"{e.Message}\n{e.StackTrace}");
            }
        }

        /// <summary>
        /// Show floating shape.
        /// </summary>
        /// <param name="sel"></param>
        public void ShowFloatingShapeSelection(Word.Selection sel)
        {
            try
            {
                // get the bytes of the shape
                var bits = (byte[])sel.EnhMetaFileBits;
                using (MemoryStream stream = new MemoryStream(bits))
                {
                    _bitmap = new Bitmap(stream);
                    lblPicSize.Content = $"{_bitmap.Width} x {_bitmap.Height}";
                    _oriWidth = _bitmap.Width;
                    _oriHeight = _bitmap.Height;
                    double ratio = (double)_oriWidth / _oriHeight;
                    double screenW = SystemParameters.PrimaryScreenWidth;
                    double screenH = SystemParameters.PrimaryScreenHeight;
                    if (_oriWidth > screenW)
                    {
                        _oriWidth = screenW;
                        _oriHeight = _oriWidth / ratio;
                    }
                    if (_oriHeight > screenH)
                    {
                        _oriHeight = screenH;
                        _oriWidth = _oriHeight * ratio;
                    }
                    _oriWidth *= 0.8;
                    _oriHeight *= 0.8;
                    uiImage.Width = _oriWidth;
                    uiImage.Height = _oriHeight;

                    // Set image source
                    using (MemoryStream ms = new MemoryStream())
                    {
                        _bitmap.Save(ms, ImageFormat.Png);
                        var bytes = ms.GetBuffer();
                        BitmapImage bi = new BitmapImage();
                        bi.BeginInit();
                        bi.StreamSource = new MemoryStream(bytes);
                        bi.EndInit();
                        uiImage.Source = bi;
                    }
                    // loading animation
                    Scale(0, 1, _oriWidth / 2, _oriHeight / 2);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show($"{e.Message}\n{e.StackTrace}");
            }
        }
        #endregion

        #region Private Methods
        private void InitUI()
        {
            // set button tooltip
            uiLogo.ToolTip = ResourceHelper.Strings.GetString("AppName");
            uiBtnSave.ToolTip = ResourceHelper.Strings.GetString("SaveAs");
            uiBtnOpenWith.ToolTip = ResourceHelper.Strings.GetString("OpenWith");
            uiBtnZoomIn.ToolTip = ResourceHelper.Strings.GetString("ZoomIn");
            uiBtnZoomOut.ToolTip = ResourceHelper.Strings.GetString("ZoomOut");
            uiBtnCentered.ToolTip = ResourceHelper.Strings.GetString("AlignCenter");
        }

        private void ZoomIn()
        {
            double scale = (double)_scale / 100;
            // scale step increase when zoom in
            if (_isZoomIn && _scale > 100) _scaleStep++;

            _scale += _scaleStep;
            _isZoomIn = true;
            double newScale = (double)_scale / 100;

            Scale(scale, newScale, _oriWidth / 2, _oriHeight / 2);

            if (!uiBtnZoomOut.IsEnabled)
            {
                uiBtnZoomOut.IsEnabled = true;
            }
        }

        private void ZoomOut()
        {
            double scale = (double)_scale / 100;
            // scale step decrease when zoom out
            if (!_isZoomIn && _scale > 100) _scaleStep--;

            if (_scale - _scaleStep < MINIMUM_SCALE) return;
            _scale -= _scaleStep;
            _isZoomIn = false;
            double newScale = (double)_scale / 100;

            Scale(scale, newScale, _oriWidth / 2, _oriHeight / 2);

            if(_scale <= MINIMUM_SCALE)
            {
                uiBtnZoomOut.IsEnabled = false;
            }
        }

        private void Scale(double from, double to, double centerX, double centerY)
        {
            uiScale.CenterX = centerX;
            uiScale.CenterY = centerY;
            // smooth scale animation
            DoubleAnimation aniScale = new DoubleAnimation() { Duration = TimeSpan.FromMilliseconds(250) };
            aniScale.From = from;
            aniScale.To = to;
            uiScale.BeginAnimation(ScaleTransform.ScaleXProperty, aniScale, HandoffBehavior.Compose);
            uiScale.BeginAnimation(ScaleTransform.ScaleYProperty, aniScale, HandoffBehavior.Compose);
            uiScaleRatio.Content = $"{_scale} %";
        }

        private System.Drawing.Rectangle GetBitmapRect(Bitmap srcBmp, Word.Shape shape)
        {
            double height = srcBmp.Height * shape.Height / (shape.Height + Math.Abs(shape.Top));
            double initW = srcBmp.Height * (shape.Width / shape.Height);
            double initH = srcBmp.Height;
            double x = 0;
            if (shape.Left > 0)
            {
                x = shape.Left * srcBmp.Height / shape.Height;
            }
            double y = 0;
            if (shape.Top > 0)
            {
                y = shape.Top * srcBmp.Height / shape.Height;
            }
            return System.Drawing.Rectangle.Empty;
        }

        /// <summary>
        /// Crop the bitmap to the specified size.
        /// </summary>
        /// <param name="srcBmp">The original bitmap.</param>
        /// <param name="width">The specified width to crop.</param>
        /// <param name="height">The specified height to crop.</param>
        /// <returns>The new bitmap.</returns>
        private Bitmap Crop(Bitmap srcBmp, int x, int y, int width, int height)
        {
            System.Drawing.Rectangle cropRect = new System.Drawing.Rectangle(x, y, width, height);
            Bitmap cropBmp = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(cropBmp))
            {
                g.DrawImage(srcBmp, new System.Drawing.Rectangle(0, 0, cropBmp.Width, cropBmp.Height),
                    cropRect, GraphicsUnit.Pixel);
            }
            return cropBmp;
        }

        private ImageFormat GetImageFormat(string extension)
        {
            extension = extension.ToLower();
            switch (extension)
            {
                case ".bmp":
                    return ImageFormat.Bmp;
                case ".jpg":
                case ".jpeg":
                    return ImageFormat.Jpeg;
                case ".png":
                    return ImageFormat.Png;
                case ".gif":
                    return ImageFormat.Gif;
                case ".emf":
                    return ImageFormat.Emf;
                case ".wmf":
                    return ImageFormat.Wmf;
                default: 
                    return ImageFormat.Bmp;
            }
        }
        #endregion

        #region Event Handler
        private void PictureViewer_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (e.Delta > 0)
            {
                ZoomIn();
            }
            else
            {
                ZoomOut();
            }
        }

        private void PictureViewer_MouseUp(object sender, MouseButtonEventArgs e)
        {
            _isPressed = false;
            _lastXOffset = uiTranslate.X;
            _lastYOffset = uiTranslate.Y;
            uiContent.Cursor = Cursors.Arrow;
        }

        private void PictureViewer_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if(e.ChangedButton == MouseButton.Left)
            {
                _isPressed = true;
                _pressedPoint = Mouse.GetPosition(uiContent);
                uiContent.Cursor = Cursors.SizeAll;
            }
        }

        private void PictureViewer_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isPressed) return;
            var mouse = Mouse.GetPosition(uiContent);
            var offset = mouse - _pressedPoint;

            offset.X /= ((double)_scale / 100);
            offset.Y /= ((double)_scale / 100);

            uiTranslate.X = _lastXOffset + offset.X;
            uiTranslate.Y = _lastYOffset + offset.Y;
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_bitmap == null)
                {
                    MessageBox.Show("Invalid Picture!");
                    return;
                }
                Microsoft.Win32.SaveFileDialog save = new Microsoft.Win32.SaveFileDialog();
                save.Filter = "Jpeg Files (*.jpg, *.jpeg)|*.jpg;*.jpeg | Png files (*.png) |*.png | Bmp Files (*.bmp)|*.bmp " +
                    "| Gif Files (*.gif)|*.gif | Emf Files (*.emf)|*.emf | All Files (*.*)|*.*";
                if (save.ShowDialog() != true) return;
                string ext = new FileInfo(save.FileName).Extension;
                ImageFormat format = GetImageFormat(ext);

                _bitmap.Save(save.FileName, format);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnOpenWith_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_bitmap == null)
                {
                    MessageBox.Show("Invalid Picture!");
                    return;
                }
                string tempFile = System.IO.Path.Combine(System.IO.Path.GetTempPath(), $"WordPictureViewer-{Guid.NewGuid()}.jpg");
                _bitmap.Save(tempFile, ImageFormat.Jpeg);
                System.Diagnostics.Process.Start(tempFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            if(_bitmap != null) _bitmap.Dispose();
            Close();
        }

        private void btnZoomOut_Click(object sender, RoutedEventArgs e)
        {
            ZoomOut();
        }

        private void btnZoomIn_Click(object sender, RoutedEventArgs e)
        {
            ZoomIn();
        }

        private void btnCentered_Click(object sender, RoutedEventArgs e)
        {
            // reset translate offset
            uiTranslate.X = 0;
            uiTranslate.Y = 0;
            _lastXOffset = 0;
            _lastYOffset = 0;
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Escape) 
            {
                if (_bitmap != null) _bitmap.Dispose();
                Close();
            }
        }



        #endregion
    }
}
