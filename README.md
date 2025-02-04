# Excelmen
Export to excel file with .NET like never before


# Usage

Use class to represent data like below and attribute above the property that persent columns

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

Use the generator

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


