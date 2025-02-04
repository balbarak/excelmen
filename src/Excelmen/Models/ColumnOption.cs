using Excelmen.Attributes;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Reflection;
using System.Text;

namespace Excelmen.Models
{
    internal class ColumnOption
    {
        public PropertyInfo Property { get; set; }

        public int Index { get; set; }

        public Type Type { get; set; }

        public bool IsRequired { get; set; }

        public string Name { get; set; }

        public string Format { get; set; }

        public bool IsAutoFitColumns { get; set; }

        public ExcelHorizontalAlignment HorizontalAlignment { get; set; }

        public ColumnOption(PropertyInfo property)
        {
            var info = property.GetCustomAttribute<ExcelColumnAttribute>();

            Index = info.Index;
            Format = info.Format;
            HorizontalAlignment = info.HorizontalAlignment;
            IsAutoFitColumns = info.AutoFitColoumns;

            Property = property;
            Name = property.GetCustomAttribute<DisplayAttribute>()?.GetName() ?? property.Name;
            Type = property.PropertyType;
        }

        public object? GetValue(object item)
        {
            return Property.GetValue(item);
        }
    }
}
