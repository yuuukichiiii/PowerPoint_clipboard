using System.Drawing;
using System.Windows.Forms;

namespace ShapePalette
{
    /// <summary>System.Drawing.Image を リボン用の stdole.IPictureDisp に変換する。</summary>
    internal sealed class PictureConverter : AxHost
    {
        private PictureConverter() : base(string.Empty) { }

        public static stdole.IPictureDisp ToPictureDisp(Image image)
        {
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }
    }
}
