# Spreadsheet Builder

[![NuGet version (SoftCircuits.SpreadsheetBuilder)](https://img.shields.io/nuget/v/SoftCircuits.SpreadsheetBuilder.svg?style=flat-square)](https://www.nuget.org/packages/SoftCircuits.SpreadsheetBuilder/)

```
Install-Package SoftCircuits.SpreadsheetBuilder
```

## Overview

SpreadsheetBuilder is a lightweight class that makes it easy to create Microsoft Excel spreadsheet (XLSX) files without Excel.

The library forgoes some features in order to keep things simple. But should be sufficient for most requirements for building an Excel spreadsheet.

The following example creates a new Excel spreadsheet file, sets a value at cell *A1*, and then saves the file.

```cs
using SpreadsheetBuilder builder = SpreadsheetBuilder.Create(Filename);
builder.SetCell("A1", "Hello, World!");
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
builder.SetCell("A1", "Hello, World!");
```

This method has dozens of overloads. You can pass a string, as shown in the example above, or you can pass other data types such as integers, doubles and decimals.

```cs
builder.SetCell("B5", 123.45);
```

The first argument specifies the cell address. You can pass the address as a string, as shown above, or you can pass an instance of the `CellReference` class. When you pass a string, the syntax is more compact but the library must convert the address to a `CellReference`. So some performance gains might be possible by passing a `CellReference` directly.

```cs
builder.SetCell(new CellReference(2, 5), 123.45);
```

To create a calculated cell, you can pass an instance of the `CellFormula` class.

```cs
builder.SetCell("C17", new CellFormula("SUM(A1:C16)");
```

## Formatting Cells

One of the more onerous tasks of building spreadsheets is creating and tracking cell formats.

Spreadsheet builder simplifies things somewhat by defining a number of predefined cell formats via the `CellStyles` property. Overloads of `SetCell()` accept a style ID parameter.

```cs
builder.SetCell("A7", 123.45, builder.CellStyles[StandardCellStyle.Currency]);
```

*Note: If you pass a `decimal` type to `SetCell()`, or typecast the value to a `decimal`, the library automatically uses the currency style when no style is specified.*

If you need something other than one of the default cell formats, you can create your own as shown in the following example.

```cs
uint bold = builder.CellStyles.Register(new CellFormat()
{
    FontId = builder.FontStyles[StandardFontStyle.Bold],
    ApplyFont = BooleanValue.FromBoolean(true),
});

uint header = builder.CellStyles.Register(new CellFormat()
{
    FontId = builder.FontStyles[StandardFontStyle.Header],
    ApplyFont = BooleanValue.FromBoolean(true)
});

uint subheader = builder.CellStyles.Register(new CellFormat()
{
    FontId = builder.FontStyles[StandardFontStyle.Subheader],
    ApplyFont = BooleanValue.FromBoolean(true)
});

uint headerRight = builder.CellStyles.Register(new CellFormat()
{
    FontId = builder.FontStyles[StandardFontStyle.Header],
    ApplyFont = BooleanValue.FromBoolean(true),
    Alignment = new() { Horizontal = HorizontalAlignmentValues.Right },
    ApplyAlignment = BooleanValue.FromBoolean(true)
});

uint subheaderRight = builder.CellStyles.Register(new CellFormat()
{
    FontId = builder.FontStyles[StandardFontStyle.Subheader],
    ApplyFont = BooleanValue.FromBoolean(true),
    Alignment = new() { Horizontal = HorizontalAlignmentValues.Right },
    ApplyAlignment = BooleanValue.FromBoolean(true)
});

builder.SetCell("D22", "Header", header);
```

In addition to the `CellStyles` property, the `SpreadsheetBuilder` class also has `NumberFormats`, `FontStyles`, `FillStyles` and `BorderStyles` properties that provide standard styles and the ability to add new ones similar to the `CellStyles` property.

*Note: When you register a style, it is stored within the current instance of `SpreadsheetBuilder`. Ensure you don't create the same styles more than once for the same instance.*

## Tables

You can create tabular data by setting the value of the appropriate cells, or you can use the `TableBuilder` class.

The `TableBuilder` class simplifies the process of creating tabular data, offers some performance gains, and can also be used to create and format a named Excel table.

The `TableBuilder` constructor takes an instance of the `SpreadsheetBuilder` class, a cell reference to the cell at the top, left corner of the table, and either of:

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

The `TableBuilder` class has a number of property for returning things like the range of the table so far.

Once you've finished building the tabular data, you can create an Excel table and style it.

```cs
table.BuilderTable("MyTableName", ExcelTableStyle.MediumBlue6);
```

## Saving

Once you've created the Excel file, you can save it to disk by calling the `Save()` or `SaveAs()` method.
