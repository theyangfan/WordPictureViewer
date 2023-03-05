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

namespace WordPictureViewer
{
    /// <summary>
    /// UserControl1.xaml 的交互逻辑
    /// </summary>
    public partial class PictureViewer : Window
    {
        #region Private Members
        private int _scale = 100;
        private int _scaleStep = 9;
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
            
            // Mouse wheel event
            MouseWheel += PictureViewer_MouseWheel;
        }
        #endregion

        public void SetImageRange(Word.Range range)
        {
            var shape = range.InlineShapes?[1];
            if (shape == null) return;
            var bits = (byte[])shape.Range.EnhMetaFileBits;
            using (MemoryStream stream = new MemoryStream(bits))
            {
                Bitmap bitmap = new Bitmap(stream);
            
                // Init image size
                double initW = bitmap.Height * (shape.Width / shape.Height);
                double initH = bitmap.Height;
                Bitmap cropBmp = Crop(bitmap, (int)initW, (int)initH);
                double ratio = (double)initW / initH;
                double screenW = SystemParameters.PrimaryScreenWidth;
                double screenH = SystemParameters.PrimaryScreenHeight;
                if (initW > screenW)
                {
                    initW = screenW;
                    initH = initW / ratio;
                }
                if (initH > screenH)
                {
                    initH = screenH;
                    initW = initH * ratio;
                }
                initW *= 0.8;
                initH *= 0.8;
                UIImage.Width = initW;
                UIImage.Height = initH;
                
                // Set image source
                using (MemoryStream ms = new MemoryStream())
                {
                    cropBmp.Save(ms, ImageFormat.Bmp);
                    byte[] bytes = ms.GetBuffer();
                    BitmapImage bi = new BitmapImage();
                    bi.BeginInit();
                    bi.StreamSource = new MemoryStream(bytes);
                    bi.EndInit();
                    UIImage.Source = bi;
                }
                double centerX = initW / 2;
                double centerY = initH / 2;
                DoubleAnimation aniScale = new DoubleAnimation() { Duration = TimeSpan.FromMilliseconds(250) };
                DoubleAnimation aniX = new DoubleAnimation(centerX, centerX, TimeSpan.FromMilliseconds(0));
                DoubleAnimation aniY = new DoubleAnimation(centerY, centerY, TimeSpan.FromMilliseconds(0));
                aniScale.From = 0;
                aniScale.To = 1;
                UIScale.BeginAnimation(ScaleTransform.ScaleXProperty, aniScale, HandoffBehavior.Compose);
                UIScale.BeginAnimation(ScaleTransform.ScaleYProperty, aniScale, HandoffBehavior.Compose);
                UIScale.BeginAnimation(ScaleTransform.CenterXProperty, aniX);
                UIScale.BeginAnimation(ScaleTransform.CenterYProperty, aniY);
            }
        }

        #region Private Methods
        private void PictureViewer_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if(e.Delta > 0) // zoom in
            {
                if(_scaleStep < 0)_scaleStep = -_scaleStep;
                else _scaleStep++;
            }
            else // zoom out
            {
                if (_scaleStep > 0) _scaleStep = -_scaleStep;
                else if(_scale + _scaleStep >= 99) _scaleStep++;
            }
            if (_scale + _scaleStep < 100) return;
            // smooth scale animation
            DoubleAnimation aniScale = new DoubleAnimation() { Duration = TimeSpan.FromMilliseconds(250) };
            _scale += _scaleStep;
            aniScale.To = (double)_scale / 100;
            System.Windows.Point mouse = Mouse.GetPosition(UIImage);
            double centerX = mouse.X;
            double centerY = mouse.Y;

            DoubleAnimation aniX = new DoubleAnimation(centerX, centerX, TimeSpan.FromMilliseconds(0));
            DoubleAnimation aniY = new DoubleAnimation(centerY, centerY, TimeSpan.FromMilliseconds(0));

            UIScale.BeginAnimation(ScaleTransform.ScaleXProperty, aniScale, HandoffBehavior.Compose);
            UIScale.BeginAnimation(ScaleTransform.ScaleYProperty, aniScale, HandoffBehavior.Compose);
            UIScale.BeginAnimation(ScaleTransform.CenterXProperty, aniX);
            UIScale.BeginAnimation(ScaleTransform.CenterYProperty, aniY);
            UIScaleRatio.Content = $"{_scale} %";
        }

        /// <summary>
        /// Crop the bitmap to the specified size.
        /// </summary>
        /// <param name="srcBmp">The original bitmap.</param>
        /// <param name="width">The specified width to crop.</param>
        /// <param name="height">The specified height to crop.</param>
        /// <returns>The new bitmap.</returns>
        private Bitmap Crop(Bitmap srcBmp, int width, int height)
        {
            System.Drawing.Rectangle cropRect = new System.Drawing.Rectangle(0, 0, width, height);
            Bitmap cropBmp = new Bitmap(width, height);
            using(Graphics g = Graphics.FromImage(cropBmp))
            {
                g.DrawImage(srcBmp, new System.Drawing.Rectangle(0, 0, cropBmp.Width, cropBmp.Height),
                    cropRect, GraphicsUnit.Pixel);
            }
            return cropBmp;
        }

        private void UICloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
        #endregion
    }
}
