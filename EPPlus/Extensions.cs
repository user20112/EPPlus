using SkiaSharp;

namespace EPPlus
{
    public static class Extensions
    {
        public static bool IsEmpty(this SKColor color)
        {
            if (color.Blue == 0 && color.Red == 0 && color.Green == 0 && color.Alpha == 0)
                return true;
            return false;
        }
    }
}