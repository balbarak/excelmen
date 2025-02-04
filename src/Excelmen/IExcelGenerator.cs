using Excelmen.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace Excelmen
{
    public interface IExcelGenerator<TModel> where TModel : class, new()
    {
        ExcelGenerator<TModel> AddRows(params TModel[] rows);
        byte[] Generate(ExcelGenerateOptions options = null);
    }
}
