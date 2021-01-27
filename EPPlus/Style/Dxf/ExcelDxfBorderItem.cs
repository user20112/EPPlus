namespace OfficeOpenXml.Style.Dxf
{
    public class ExcelDxfBorderItem : DxfStyleBase<ExcelDxfBorderItem>
    {
        internal ExcelDxfBorderItem(ExcelStyles styles) :
            base(styles)
        {
            Color=new ExcelDxfColor(styles);
        }

        public ExcelDxfColor Color { get; internal set; }
        public ExcelBorderStyle? Style { get; set; }

        protected internal override bool HasValue
        {
            get
            {
                return Style != null || Color.HasValue;
            }
        }

        protected internal override string Id
        {
            get
            {
                return GetAsString(Style) + "|" + (Color == null ? "" : Color.Id);
            }
        }

        protected internal override ExcelDxfBorderItem Clone()
        {
            return new ExcelDxfBorderItem(_styles) { Style = Style, Color = Color };
        }

        protected internal override void CreateNodes(XmlHelper helper, string path)
        {
            SetValueEnum(helper, path + "/@style", Style);
            SetValueColor(helper, path + "/d:color", Color);
        }
    }
}