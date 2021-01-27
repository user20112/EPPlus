using SkiaSharp;

namespace OfficeOpenXml.Compatibility
{
    internal class ImageCompat
    {
        internal static byte[] GetImageAsByteArray(SKImage image)
        {
            if (image.EncodedData != null)
                return image.EncodedData.ToArray();
            return image.Encode().ToArray();
        }
    }
}