using System;
using System.Collections.Generic;
using System.Linq;
using System.Resources;
using System.Text;
using System.Threading.Tasks;

namespace WordPictureViewer
{
    internal class ResourceHelper
    {
        private static ResourceManager _resourceManager;

        public static ResourceManager Current
        {
            get
            {
                if(_resourceManager == null)
                {
                    _resourceManager = new ResourceManager("WordPictureViewer.Resources.Strings", typeof(ResourceHelper).Assembly);
                }
                return _resourceManager;
            }
        }
    }
}
