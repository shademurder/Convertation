using System;

namespace CaseAgile.OfficePublisher
{
    class Parameter
    {
        private readonly String value;

        /// <summary>
        /// Max rows in excel sheet
        /// </summary>
        public static readonly Parameter ExcelMaxRows = new Parameter("ExcelMaxRows");
        /// <summary>
        /// Max columns in excel sheet
        /// </summary>
        public static readonly Parameter ExcelMaxColumns = new Parameter("ExcelMaxColumns");
        /// <summary>
        /// Max rows and columns in excel sheet
        /// </summary>
        public static readonly Parameter ExcelMaxSheetSize = new Parameter("ExcelMaxSheetSize");

        private Parameter(String value)
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
