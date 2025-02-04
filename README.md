# Excelmen
Export to excel file with .NET like never before


# Usage

Create `class` to represent the data like below

```
public class Person
{
    [ExcelColumn(Index = 1)]
    public string FirstName { get; set; }

    [ExcelColumn(Index = 2)]
    public string LastName { get; set; }

    public Person(string firstName,string lastName)
    {
        FirstName = firstName;
        LastName = lastName;
    }
}
```

Use the `ExcelGenerator` to export person data to excel

```
var rows = new Person[]
{
    new Person("Faisal","Ahmed")
};

var generator = new ExcelGenerator<Person>();

generator.AddRows(rows);

var excelBytes = generator.Generate();

File.WriteAllBytes(@"Path\to\exportext.xlsx",excelBytes)

```


