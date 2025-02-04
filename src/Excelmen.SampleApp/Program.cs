using Excelmen.Attributes;
using OfficeOpenXml.Style;
using System.ComponentModel.DataAnnotations;

namespace Excelmen.SampleApp
{
    internal class Program
    {

        static void Main(string[] args)
        {
            var generator = new ExcelGenerator<Person>();

            var data = new Person[]
            {
                new Person()
                {
                    FirstName = "Khalid",
                    LastName = "Mohammad",
                    Number = 20044
                }
            };

            generator.AddRows(data);

            var excelBytes = generator.Generate();

            File.WriteAllBytes(@"C:\Users\balba\Desktop\excels\sample.xlsx", excelBytes);
        }
    }

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
            Format = "dd-MMM-yyyy hh:mm:ss",
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
