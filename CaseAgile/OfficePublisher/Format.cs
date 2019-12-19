using System;

namespace CaseAgile.OfficePublisher
{
    class Format
    {
        private readonly String value;

        public static readonly Format Doc = new Format(".doc");
        public static readonly Format Docx = new Format(".docx");
        public static readonly Format Pdf = new Format(".pdf");
        public static readonly Format Ppt = new Format(".ppt");
        public static readonly Format Pptx = new Format(".pptx");
        public static readonly Format Xls = new Format(".xls");
        public static readonly Format Xlsx = new Format(".xlsx");

        private Format(String value)
        {
            this.value = value;
        }

        public override String ToString()
        {
            return value;
        }

        public String Value => value;
    }
}
