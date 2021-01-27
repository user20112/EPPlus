using SkiaSharp;
using System;

namespace OfficeOpenXml.Style.Dxf
{
    public class ExcelDxfColor : DxfStyleBase<ExcelDxfColor>

    {
        public ExcelDxfColor(ExcelStyles styles) : base(styles)
        {
        }

        public bool? Auto { get; set; }
        public SKColor? Color { get; set; }
        public int? Index { get; set; }
        public int? Theme { get; set; }
        public double? Tint { get; set; }

        protected internal override bool HasValue
        {
            get
            {
                return Theme != null ||
                       Index != null ||
                       Auto != null ||
                       Tint != null ||
                       Color != null;
            }
        }

        protected internal override string Id
        {
            get { return GetAsString(Theme) + "|" + GetAsString(Index) + "|" + GetAsString(Auto) + "|" + GetAsString(Tint) + "|" + GetAsString(Color==null ? "" : ((uint)((SKColor)Color.Value)).ToString("x")); }
        }

        protected internal override ExcelDxfColor Clone()
        {
            return new ExcelDxfColor(_styles) { Theme = Theme, Index = Index, Color = Color, Auto = Auto, Tint = Tint };
        }

        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            throw new NotImplementedException();
        }
    }
}