using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Resources;
using System.Text;
using System.Threading.Tasks;

namespace WordPictureViewer
{
    internal class ResourceHelper
    {
        private static ResourceManager _resourceManager;

        /// <summary>
        /// The localizable strings resource.
        /// </summary>
        public static ResourceManager Strings
        {
            get
            {
                if (_resourceManager == null)
                {
                    _resourceManager = new ResourceManager("WordPictureViewer.Resources.Strings", typeof(ResourceHelper).Assembly);
                }
                return _resourceManager;
            }
        }
    }
}
