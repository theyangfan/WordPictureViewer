﻿using System;
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
        private const int MINIMUM_SCALE = 50;
        private int _scale = 100;
        private int _scaleStep = 10;
        private bool _zoomIn = true;
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
                    Bitmap bitmap = new Bitmap(stream);
                    // get the available width
                    double initW = bitmap.Height * (shape.Width / shape.Height);
                    double initH = bitmap.Height;
                    // crop the available area
                    Bitmap cropBmp = Crop(bitmap, 0, 0, (int)initW, (int)initH);
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

                    // set image source
                    using (MemoryStream ms = new MemoryStream())
                    {
                        cropBmp.Save(ms, ImageFormat.Png);
                        byte[] bytes = ms.GetBuffer();
                        BitmapImage bi = new BitmapImage();
                        bi.BeginInit();
                        bi.StreamSource = new MemoryStream(bytes);
                        bi.EndInit();
                        UIImage.Source = bi;
                    }
                    // loading animation
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
                    Bitmap bitmap = new Bitmap(stream);
                    double initW = bitmap.Width;
                    double initH = bitmap.Height;
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
                        bitmap.Save(ms, ImageFormat.Png);
                        byte[] bytes = ms.GetBuffer();
                        BitmapImage bi = new BitmapImage();
                        bi.BeginInit();
                        bi.StreamSource = new MemoryStream(bytes);
                        bi.EndInit();
                        UIImage.Source = bi;
                    }
                    // loading animation
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
            catch (Exception e)
            {
                MessageBox.Show($"{e.Message}\n{e.StackTrace}");
            }
        }
        #endregion

        #region Private Methods

        private void PictureViewer_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if(e.Delta > 0) // zoom in
            {
                if(_zoomIn && _scale > 100) _scaleStep++;
                _scale += _scaleStep;
                _zoomIn = true;
            }
            else // zoom out
            {
                if (!_zoomIn && _scale > 100) _scaleStep--;
                if (_scale - _scaleStep < MINIMUM_SCALE) return;
                _scale -= _scaleStep;
                _zoomIn = false;
            }
            // smooth scale animation
            DoubleAnimation aniScale = new DoubleAnimation() { Duration = TimeSpan.FromMilliseconds(250) };
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

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.Key == Key.Escape) 
            {
                Close();
            }
        }
        #endregion



    }
}
