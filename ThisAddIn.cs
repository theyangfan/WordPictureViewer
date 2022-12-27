using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows;
using F = System.Windows.Forms;
using C = System.Windows.Controls;
using D = System.Drawing;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace WordPictureViewer
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.WindowBeforeDoubleClick += Application_WindowBeforeDoubleClick;
        }

        private void Application_WindowBeforeDoubleClick(Word.Selection Sel, ref bool Cancel)
        {
            if(Sel.Type == Word.WdSelectionType.wdSelectionShape
                || Sel.Type == Word.WdSelectionType.wdSelectionInlineShape)
            {
                try
                {
                    Window wnd = new Window();
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
                    wnd.Show();
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message+"\n"+e.StackTrace);
                }
                
            }
            
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
