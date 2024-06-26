﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows;
using System.Drawing;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Globalization;

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
            Ribbon ribbon = Globals.Ribbons.GetRibbon<Ribbon>();
            if (!ribbon.IsEnable) return;
            // inline shapes
            if(Sel.Type == Word.WdSelectionType.wdSelectionInlineShape)
            {
                PictureViewer viewer = new PictureViewer();
                viewer.Show();
                viewer.ShowInlineShapeSelection(Sel);
            }
            // floating shapes
            else if(Sel.Type == Word.WdSelectionType.wdSelectionShape)
            {
                PictureViewer viewer = new PictureViewer();
                viewer.Show();
                viewer.ShowFloatingShapeSelection(Sel);
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
