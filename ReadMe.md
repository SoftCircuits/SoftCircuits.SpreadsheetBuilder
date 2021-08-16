# Spreadsheet Builder

[![NuGet version (SoftCircuits.SpreadsheetBuilder)](https://img.shields.io/nuget/v/SoftCircuits.SpreadsheetBuilder.svg?style=flat-square)](https://www.nuget.org/packages/SoftCircuits.SpreadsheetBuilder/)

```
Install-Package SoftCircuits.SpreadsheetBuilder
```

## Overview

SpreadsheetBuilder is a lightweight class that makes it easy to create Microsoft Excel spreadsheet files (XLSX) without Excel.

The library gives up some features to keep things simple. The following example creates a new Excel spreadsheet file, sets a value at cell *A1*, and then saves the file.

```cs
using SpreadsheetBuilder builder = SpreadsheetBuilder.Create(Filename);
builder.SetCell("A1", "Hello, World!);
builder.Save();
```
## Getting Started

To get started, create an instance of the `SpreadsheetBuilder` class. You can do that using any of the following static methods of the `SpreadsheetBuilder` class.

```cs
public static SpreadsheetBuilder Create(string path, SpreadsheetDocumentType type = SpreadsheetDocumentType.Workbook);

public static SpreadsheetBuilder CreateFromTemplate(string path);

public static SpreadsheetBuilder Open(string path, bool isEditable);
```

The `SpreadsheetBuilder` class implements `IDisposable`, so you should use a `using` statement to ensure the class cleans up in a timely manner.

```cs
using SpreadsheetBuilder builder = SpreadsheetBuilder.Create(Filename);
```

## Setting Cell Values

To set the value of a cell, use the `SetCell()` method.

```cs
builder.SetCell("A1", "Hello, World!);
```

This method has dozens of overloads. You can pass a string, as shown in the example above, or you can pass other data types such as integers, doubles and decimals.

```cs
builder.SetCell("B5", 123.45);
```

The first argument specifies the cell address. You can pass the address as a string, as shown above, or you can pass an instance of the `CellReference` class. When you pass a string, the syntax is more compact but the library must convert the address to a `CellReference`. So some performance gains might be possible by passing a `CellReference` directly.

```cs
builder.SetCell(new CellReference(2, 5), 123.45);
```

> TODO: Formulas

## Formatting Cells

> TODO: styleId parameter
> TODO: styles

## Tables

You can create tabular data by setting the value of the appropriate cells, or you can use the `TableBuilder` class.

The `TableBuilder` class, simplifies the process of creating tabular data, offers some performance gains, and can also be used to create and formatting a named Excel table.

The constructor takes an instance of the `SpreadsheetBuilder` class, a cell reference to the cell at the top, left corner of the table, and either of:

- The number of columns
- An `IEnumerable<string>` of the column headers

If you specify the number of columns, it is assumed the table has no headers.

To write data to the table, call the `AddRow()` method. This method accepts any number of arguments, each of which is assigned to the corresponding cell on the current table row. The type of the arguments can be `string`, `int`, `double`, etc. They can also be an instance of `CellValue<T>`, which can specify a style ID in addition to a value. In addition, they can also be an instance of `CellFormula`.

```cs
string[] columns = new string[]
{
  "Column1",
  "Column2",
  "Column3"
};

TableBuilder table = new(builder, "A4", columns);
table.AddRow(new CellValue<string>("Abc", 123, 123.45);
table.AddRow(new CellValue<string>("Def", 456, (decimal)4000);
```

Once you've finished building the tabular data, you can create an Excel table and style it.

```cs
table.BuilderTable("MyTableName", ExcelTableStyle.MediumBlue6);
```

