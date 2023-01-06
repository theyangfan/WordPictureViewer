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

namespace WordPictureViewer
{
    /// <summary>
    /// UserControl1.xaml 的交互逻辑
    /// </summary>
    public partial class PictureViewer : Window
    {
        private int _scale = 100;
        private int _scaleStep = 9;
        public PictureViewer(Bitmap bitmap)
        {
            InitializeComponent();
            ResizeMode = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            WindowState = WindowState.Maximized;
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;

            double screenW = SystemParameters.PrimaryScreenWidth;
            double screenH = SystemParameters.PrimaryScreenHeight - 35;

            double initW = bitmap.Width;
            double initH = bitmap.Height;
            double ratio = (double)bitmap.Width / bitmap.Height;

            if(initW > screenW)
            {
                initW= screenW;
                initH = initW / ratio;
            }
            if(initH > screenH)
            {
                initH = screenH;
                initW = initH * ratio;
            }
            initW *= 0.8;
            initH *= 0.8;
            UIImage.Width = initW;
            UIImage.Height = initH;
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
            
            MouseWheel += PictureViewer_MouseWheel;
        }

        private void PictureViewer_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if(e.Delta > 0)
            {
                if(_scaleStep < 0)_scaleStep = -_scaleStep;
                else _scaleStep++;
            }
            else
            {
                if (_scaleStep > 0) _scaleStep = -_scaleStep;
                else if(_scale + _scaleStep >= 99) _scaleStep++;
            }
            
            if (_scale + _scaleStep < 100) return;
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
            

            //UIImage.RenderTransform = new ScaleTransform(scale, scale, centerX, centerY);
            UIScaleRatio.Content = $"{_scale} %";
        }

        private void UICloseBtn_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
