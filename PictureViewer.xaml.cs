using System;
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
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WordPictureViewer
{
    /// <summary>
    /// UserControl1.xaml 的交互逻辑
    /// </summary>
    public partial class PictureViewer : Window
    {
        public PictureViewer()
        {
            InitializeComponent();
            ResizeMode = ResizeMode.NoResize;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            WindowState = WindowState.Maximized;
            WindowStyle = WindowStyle.None;
            AllowsTransparency = true;
            /*Window wnd = new Window();
                    wnd.ResizeMode = ResizeMode.NoResize;
                    wnd.WindowStartupLocation = WindowStartupLocation.CenterOwner;
                    wnd.WindowState = WindowState.Maximized;
                    wnd.WindowStyle = WindowStyle.None;
                    wnd.AllowsTransparency = true;
                    wnd.Background = new SolidColorBrush(Color.FromArgb(200, 0, 0, 0));

                    C.StackPanel panel = new C.StackPanel();
                    C.Button close = new C.Button() { Content = "关闭" };
                    close.Click += (sen, e) => {
                        wnd.Close();
                    };
                    close.HorizontalAlignment = HorizontalAlignment.Right;
                    C.Image image = new C.Image();
                    image.Stretch = Stretch.Uniform;
                    image.HorizontalAlignment = HorizontalAlignment.Center;
                    C.Label label = new C.Label();
                    label.Foreground = new SolidColorBrush(Color.FromRgb(255, 255, 255));
                    var bits = (byte[])Sel.Range.EnhMetaFileBits;
                    using (MemoryStream ms = new MemoryStream(bits), ms2 = new MemoryStream())
                    {
                        D.Bitmap bitmap = new D.Bitmap(ms);
                        bitmap.Save(ms2, D.Imaging.ImageFormat.Png);
                        byte[] bytes = ms2.GetBuffer();
                        BitmapImage bi = new BitmapImage();
                        bi.BeginInit();
                        bi.StreamSource = new MemoryStream(bytes);
                        bi.EndInit();
                        image.Source = bi;
                    }
                    panel.Children.Add(close);
                    panel.Children.Add(label);
                    panel.Children.Add(image);

                    double size = SystemParameters.PrimaryScreenWidth;
                    wnd.MouseWheel += (sen, e) =>
                    {
                        size += e.Delta;
                        image.Width = size;
                        label.Content = size;
                    };

                    wnd.Content = panel;
                    wnd.Show();*/
        }
    }
}
