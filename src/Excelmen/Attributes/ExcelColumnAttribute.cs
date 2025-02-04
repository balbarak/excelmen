using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excelmen.Attributes
{
    public class ExcelColumnAttribute : Attribute
    {
        public int Index { get; set; }

        public string Format { get; set; }

        public ExcelHorizontalAlignment HorizontalAlignment { get; set; }

        public bool AutoFitColoumns { get; set; } = true;
    }
}
