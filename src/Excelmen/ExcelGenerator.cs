using Excelmen.Attributes;
using Excelmen.Models;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Excelmen
{
    public class ExcelGenerator<TModel> : IExcelGenerator<TModel> where TModel : class, new()
    {
        private readonly int _startRowIndex = 2;
        private readonly ExcelPackage _package;
        private readonly ExcelWorksheet _sheet;
        private readonly List<TModel> _rows;
        private List<ColumnOption> _columns;

        private string _sheetName;

        public ExcelGenerator(string sheetName = "sheet1")
        {
            _sheetName = sheetName;
            _package = new ExcelPackage();
            _sheet = _package.Workbook.Worksheets.Add(sheetName);
            _sheetName = sheetName;
            _rows = new List<TModel>();
            _columns = new List<ColumnOption>();
        }

        public ExcelGenerator<TModel> AddRows(params TModel[] rows)
        {
            _rows.AddRange(rows);

            return this;
        }

        public byte[] Generate(ExcelGenerateOptions options = null)
        {
            if (options == null)
                options = new ExcelGenerateOptions();

            byte[] result;

            SetColumns();

            InsertRows();

            if (options.FormatAsTable)
                FormatAsTable();

            if (options.AutoFit)
                AutoFitAll();

            using (MemoryStream ms = new MemoryStream())
            {
                _package.SaveAs(ms);

                result = ms.ToArray();
            }


            return result;
        }

        private void SetColumns()
        {
            _columns = GetColumnsOptions();

            foreach (var item in _columns)
            {
                AddColumn(item);
            }
        }

        private List<ColumnOption> GetColumnsOptions()
        {
            var type = typeof(TModel);

            var result = new List<ColumnOption>();

            var properties = type.GetProperties();

            foreach (var item in properties)
            {
                var excelColumnInfo = item.GetCustomAttribute<ExcelColumnAttribute>();

                if (excelColumnInfo != null)
                    result.Add(new ColumnOption(item));
            }

            return result;
        }

        private void AddColumn(ColumnOption option)
        {
            _sheet.InsertColumn(option.Index, 1);

            var cell = _sheet.Cells[1, option.Index];

            cell.Value = option.Name;

            cell.Style.HorizontalAlignment = option.HorizontalAlignment;

            cell.AutoFitColumns();

            var fullColomAddress = GetFullColumnAddress(option.Index);
            var rows = _sheet.Cells[fullColomAddress];

            if (option.IsAutoFitColumns)
                rows.AutoFitColumns();

            rows.Style.HorizontalAlignment = option.HorizontalAlignment;

            if (!string.IsNullOrWhiteSpace(option.Format))
                rows.Style.Numberformat.Format = option.Format;
        }

        private void InsertRows()
        {
            var rowIndex = _startRowIndex;

            foreach (var item in _rows)
            {
                foreach (var column in _columns)
                {
                    object value = column.GetValue(item);

                    AddCell(rowIndex, column.Index, value, column);
                }

                rowIndex++;
            }
        }

        private void AddCell(int rowIndex, int colomIndex, object value, ColumnOption options)
        {
            var cell = _sheet.Cells[rowIndex, colomIndex];

            cell.Value = value;

            var fullColomAddress = GetFullColumnAddress(colomIndex);

            if (options.IsAutoFitColumns)
                cell.AutoFitColumns();
        }

        private string GetColumnAddress(int colomIndex)
        {
            var result = ExcelCellBase.GetAddress(1, colomIndex);
            result = result.Replace("1", "").Trim();

            return result;
        }

        private string GetFullColumnAddress(int colomIndex)
        {
            var result = GetColumnAddress(colomIndex);

            result = $"{result}:{result}";

            return result;
        }

        public virtual void FormatAsTable()
        {
            int endColumn = _sheet.Dimension.End.Column;
            int endRow = _sheet.Dimension.End.Row;

            var range = _sheet.Cells[1, 1, endRow, endColumn];
            var table = _sheet.Tables.Add(range, "table1");

            table.TableStyle = TableStyles.Medium9;
            range.AutoFitColumns();
        }

        public void AutoFitAll()
        {
            var range = _sheet.Cells[1, 1, _sheet.Dimension.End.Row, _sheet.Dimension.End.Column];
            range.AutoFitColumns();
        }
    }
}
