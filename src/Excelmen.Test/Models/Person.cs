using Excelmen.Attributes;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excelmen.Test.Models
{
    internal class Person
    {
        [ExcelColumn(Index = 1)]
        [Display(Name = "First Name")]
        public string FirstName { get; set; }

        [ExcelColumn(Index = 2)]
        [Display(Name = "Last Name")]
        public string LastName { get; set; }


        [ExcelColumn(
            Index = 3,
            Format ="dd-MMM-yyyy hh:mm:ss",
            AutoFitColoumns = true,
            HorizontalAlignment = ExcelHorizontalAlignment.Center)]
        [Display(Name = "Date of Birth")]
        public DateTime DateOfBirth { get; set; }


        [ExcelColumn(
            Index = 4,
            Format = "##,#0.00",
            HorizontalAlignment = ExcelHorizontalAlignment.Center)]
        [Display(Name = "Number")]
        public decimal Number { get; set; }
    }
}
