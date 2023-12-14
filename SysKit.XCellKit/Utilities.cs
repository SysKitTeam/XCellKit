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

        //public static double CalculateHorizontalDpi(this ImageMetadata metadata)
        //{
        //    switch (metadata.ResolutionUnits)
        //    {
        //        case PixelResolutionUnit.PixelsPerMeter:
        //            return metadata.HorizontalResolution * 0.0254;
        //        case PixelResolutionUnit.PixelsPerInch:
        //            return metadata.HorizontalResolution * 2.54;
        //        case PixelResolutionUnit.AspectRatio:
        //        case PixelResolutionUnit.PixelsPerCentimeter:
        //        default:
        //            return metadata.HorizontalResolution;
        //    }
        //}

        //public static double CalculateVerticalDpi(this ImageMetadata metadata)
        //{
        //    switch (metadata.ResolutionUnits)
        //    {
        //        case PixelResolutionUnit.PixelsPerMeter:
        //            return metadata.VerticalResolution * 0.0254;
        //        case PixelResolutionUnit.PixelsPerInch:
        //            return metadata.VerticalResolution * 2.54;
        //        case PixelResolutionUnit.AspectRatio:
        //        case PixelResolutionUnit.PixelsPerCentimeter:
        //        default:
        //            return metadata.VerticalResolution;
        //    }
        //}

        public static string GetRgbAsHex(this SKColor color)
        {
            return $"{color.Red:X2}{color.Green:X2}{color.Blue:X2}";
        }

        public static int GetRgbAsInt(this SKColor color)
        {
            return (color.Alpha << 24) | (color.Red << 16) | (color.Green << 8) | color.Blue;
        }
    }
}
