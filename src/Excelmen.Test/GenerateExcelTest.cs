using Excelmen.Test.Models;

namespace Excelmen.Test
{
    public class GenerateExcelTest
    {
        [Fact]
        public void Should_Create_Excel_File()
        {
            var data = new Person[]
            {
                new Person()
                {
                    FirstName = "Fahad",
                    LastName = "John",
                    DateOfBirth = DateTime.Now.AddYears(-22),
                    Number = 2039485
                },
                new Person()
                {
                    FirstName = "Nasser",
                    LastName = "Turki",
                    DateOfBirth = DateTime.Now.AddYears(-22),
                    Number = 2_949_493
                },
                new Person()
                {
                    FirstName = "Majed",
                    LastName= "Nicolas",
                    DateOfBirth = DateTime.Now.AddYears(-22),
                    Number = 41_499
                }
            };

            var generator = new ExcelGenerator<Person>();

            generator.AddRows(data);

            var result = generator.Generate();

            File.WriteAllBytes(@"C:\Users\balba\Desktop\excels\persons.xlsx", result);
        }
    }
}