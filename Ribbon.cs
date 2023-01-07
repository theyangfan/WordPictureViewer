using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordPictureViewer
{
    public partial class Ribbon
    {
        public bool IsEnable
        {
            get => UIEnable.Checked;
            set => UIEnable.Checked = value;
        }

        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }
    }
}
