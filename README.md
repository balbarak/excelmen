# Excelmen
Export to excel file with .NET like never before


# Installation

csproj Reference

`<PackageReference Include="Excelmen" Version="1.0.0" />`

donet CLI

`dotnet add package Excelmen --version 1.0.0`

# Usage

Create `class` to represent the data and use `ExcelColumn` to represent excel column specification

```C#
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

```C#
var rows = new Person[]
{
    new Person("Faisal","Ahmed")
};

var generator = new ExcelGenerator<Person>();

generator.AddRows(rows);

var excelBytes = generator.Generate();

File.WriteAllBytes(@"Path\to\expotFile.xlsx",excelBytes)

```


