using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace WordPictureViewer
{
    public partial class Ribbon
    {
        public bool IsEnable
        {
            get => uiEnable.Checked;
            set => uiEnable.Checked = value;
        }

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {
            uiEnable.Label = ResourceHelper.Strings.GetString("Enable");
            Version version = Assembly.GetExecutingAssembly().GetName().Version;
            uiVersion.Label = $"(V{version.Major}.{version.Minor}.{version.Build})";
        }
    }
}
