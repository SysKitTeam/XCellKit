using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SysKit.XCellKit
{
    static class Utilities
    {

        public static Uri SafelyCreateUri(string uriString, UriKind uriKind)
        {
            Uri uri = null;
            try
            {
                uri = new Uri(uriString, uriKind);
               // uri constructor will not always throw an error, so we do it in this a little bit hackish way
                var absoluteUri = uri.AbsoluteUri;
            }
            catch
            {
                uri = null;
            }
            return uri;
        }

        public static Uri SafelyCreateUri(string uriString)
        {
            Uri uri = null;
            try
            {
                uri = new Uri(uriString);
                // uri constructor will not always throw an error, so we do it in this a little bit hackish way
                var absoluteUri = uri.AbsoluteUri;
            }
            catch
            {
                uri = null;
            }
            return uri;
        }
    }
}
