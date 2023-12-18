using SkiaSharp;
using System;

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

        public static string GetRgbAsHex(this SKColor color)
        {
            return $"{color.Red:X2}{color.Green:X2}{color.Blue:X2}";
        }

        public static int GetRgbAsInt(this SKColor color)
        {
            return (color.Alpha << 24) | (color.Red << 16) | (color.Green << 8) | color.Blue;
        }

        public static string GetId(this SKFont font)
        {
            return $"[Font: Name={font.Typeface.FamilyName}, Size={font.Size}]";
        }
    }
}
