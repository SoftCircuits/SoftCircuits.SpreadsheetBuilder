// Copyright (c) 2021 Jonathan Wood (www.softcircuits.com)
// Licensed under the MIT license.
//
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace SoftCircuits.Spreadsheet
{
    /// <summary>
    /// Main Excel spreadsheet builder class.
    /// </summary>
    public partial class SpreadsheetBuilder : IDisposable
    {
        private const string DefaultSheetName = "Sheet1";

        /// <summary>
        /// Gets or sets whether the save methods (<see cref="Save"/> and <see cref="SaveAs(string)"/>)
        /// throw an exception if the current document does not validate.
        /// </summary>
        public static SaveValidation ValidationExceptions { get; set; } = SaveValidation.None;

        #region Public properties

        /// <summary>
        /// Gets the current <see cref="SpreadsheetDocument"/>.
        /// </summary>
        public SpreadsheetDocument Document { get; private set; }

        /// <summary>
        /// Gets the current document's <see cref="WorkbookPart"/>.
        /// </summary>
        public WorkbookPart WorkbookPart { get; private set; }

        /// <summary>
        /// Gets or sets the current document's active worksheet.
        /// </summary>
        public Worksheet? Worksheet { get; set; }

        /// <summary>
        /// Gets the current document's collection of numbering formats.
        /// </summary>
        public NumberFormats NumberFormats { get; private set; }

        /// <summary>
        /// Gets the current document's collection of font styles.
        /// </summary>
        public FontStyles FontStyles { get; private set; }

        /// <summary>
        /// Gets the current document's collection of fill styles.
        /// </summary>
        public FillStyles FillStyles { get; private set; }

        /// <summary>
        /// Gets the current document's collection of border styles.
        /// </summary>
        public BorderStyles BorderStyles { get; private set; }

        /// <summary>
        /// Gets the current document's collection of cell styles.
        /// </summary>
        public CellStyles CellStyles { get; private set; }

        #endregion

        #region Construction

        public static SpreadsheetBuilder Create(string path, SpreadsheetDocumentType type = SpreadsheetDocumentType.Workbook) =>
            new(SpreadsheetDocument.Create(path, type), true);
            
        public static SpreadsheetBuilder Create(System.IO.Stream stream, SpreadsheetDocumentType type = SpreadsheetDocumentType.Workbook) =>
            new(SpreadsheetDocument.Create(stream, type), true);

        public static SpreadsheetBuilder CreateFromTemplate(string path) =>
            new(SpreadsheetDocument.CreateFromTemplate(path), false);

        public static SpreadsheetBuilder Open(string path, bool isEditable) =>
            new(SpreadsheetDocument.Open(path, isEditable), false);

        /// <summary>
        /// Private constructor.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="isCreating"></param>
        private SpreadsheetBuilder(SpreadsheetDocument document, bool isCreating)
        {
            Document = document ?? throw new ArgumentNullException(nameof(document));

            // Ensure that WorkbookPart is not null and that we have all the standard elements
            if (isCreating)
            {
                WorkbookPart = Document.AddWorkbookPart();
                WorkbookPart.Workbook = new Workbook();
                WorkbookPart.Workbook.AppendChild(new Sheets());
                Worksheet = CreateWorksheet(DefaultSheetName);
            }
            else
            {
                WorkbookPart = Document.WorkbookPart ?? Document.AddWorkbookPart();
                if (WorkbookPart.Workbook == null)
                    WorkbookPart.Workbook = new Workbook();
                Sheets sheets = WorkbookPart.Workbook.Elements<Sheets>().FirstOrDefault() ??
                    WorkbookPart.Workbook.AppendChild(new Sheets());
                if (!sheets.Elements<Sheet>().Any())
                    Worksheet = CreateWorksheet(DefaultSheetName);
            }

            // When we create a stylesheet, Excel requires us to have certain
            // stylesheet elements defined, and in the correct order
            NumberFormats = new(this);
            FontStyles = new(this);
            FillStyles = new(this);
            BorderStyles = new(this);
            CellStyles = new(this);
        }

        #endregion

        #region Saving

        /// <summary>
        /// Saves the current document.
        /// </summary>
        public void Save()
        {
#if DEBUG
            if (ValidationExceptions == SaveValidation.Always || ValidationExceptions == SaveValidation.DebugOnly)
#else
            if (ValidationExceptions == SaveValidation.Always)
#endif
                ThrowExceptionOnValidationErrors(nameof(Save));

            Document.Save();
        }

        /// <summary>
        /// Saves the current document with the specified file name.
        /// </summary>
        /// <param name="path"></param>
        public void SaveAs(string path)
        {
#if DEBUG
            if (ValidationExceptions == SaveValidation.Always || ValidationExceptions == SaveValidation.DebugOnly)
#else
            if (ValidationExceptions == SaveValidation.Always)
#endif
                ThrowExceptionOnValidationErrors(nameof(Save));

            Document.SaveAs(path);
        }

        /// <summary>
        /// Performs validation on the current document and returns any validation errors.
        /// Returns an empty <see cref="IEnumerable{T}"/> if the current document is valid.
        /// </summary>
        public IEnumerable<ValidationErrorInfo> GetValidationErrors()
        {
            OpenXmlValidator validator = new();
            return validator.Validate(Document);
        }

        /// <summary>
        /// Performs validation on the current document and then throws an exception
        /// if any validation errors were detected.
        /// </summary>
        /// <param name="methodName"></param>
        private void ThrowExceptionOnValidationErrors(string methodName)
        {
            var errors = GetValidationErrors();
            if (errors.Any())
                throw new Exception($"Validation failed in {methodName}() : {string.Join(", ", errors.Select(e => e.Description))}");
        }

        #endregion

        #region Worksheets

        /// <summary>
        /// Returns the first worksheet in the workbook, or null if there are no
        /// worksheets.
        /// </summary>
        /// <returns>The first worksheet in the workbook.</returns>
        public Worksheet? GetFirstWorksheet()
        {
            Sheets? sheets = WorkbookPart.Workbook?.GetFirstChild<Sheets>();
            Sheet? sheet = sheets?.Elements<Sheet>().FirstOrDefault();
            if (sheet?.Id?.Value != null)
                return ((WorksheetPart)WorkbookPart.GetPartById(sheet.Id!)).Worksheet;
            return null;
        }

        /// <summary>
        /// Returns the worksheet with the specified name, or null if no worksheets
        /// found with that name.
        /// </summary>
        /// <param name="name">The name of the worksheet to return.</param>
        /// <returns>The matching worksheet.</returns>
        public Worksheet? GetWorksheet(string name)
        {
            Sheets? sheets = WorkbookPart.Workbook?.GetFirstChild<Sheets>();
            Sheet? sheet = sheets?.Elements<Sheet>()
                .Where(s => string.Compare(s.Name, name, true) == 0)
                .FirstOrDefault();
            if (sheet?.Id?.Value != null)
                return ((WorksheetPart)WorkbookPart.GetPartById(sheet.Id!)).Worksheet;
            return null;
        }

        /// <summary>
        /// Creates a new worksheet with the specified name.
        /// </summary>
        /// <param name="name">Name for new worksheet.</param>
        /// <returns>The newly created worksheet.</returns>
        public Worksheet CreateWorksheet(string name)
        {
            // Add a new worksheet part to the workbook
            WorksheetPart worksheetPart = WorkbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = WorkbookPart.Workbook.Elements<Sheets>().FirstOrDefault() ??
                WorkbookPart.Workbook.AppendChild(new Sheets());

            string relationshipId = WorkbookPart.GetIdOfPart(worksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId?.Value ?? 0).DefaultIfEmpty((uint)0).Max() + 1;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new() { Id = relationshipId, SheetId = sheetId, Name = name };
            sheets.Append(sheet);

            return worksheetPart.Worksheet;
        }

        /// <summary>
        /// Renames the specified worksheet.
        /// </summary>
        /// <param name="worksheet">The worksheet to be renamed.</param>
        /// <param name="name">The new name to give the worksheet.</param>
        public void RenameWorksheet(Worksheet worksheet, string name)
        {
            WorksheetPart? worksheetPart = worksheet?.WorksheetPart;
            if (worksheetPart != null)
            {
                Sheets? sheets = WorkbookPart.Workbook?.GetFirstChild<Sheets>();
                if (sheets != null)
                {
                    string id = WorkbookPart.GetIdOfPart(worksheetPart);
                    Sheet? sheet = sheets.Elements<Sheet>().FirstOrDefault(s => s.Id == id);
                    if (sheet != null)
                        sheet.Name = name;
                }
            }
        }

        #endregion

        #region Create and delete cells

        /// <summary>
        /// Returns the specified cell, creating it if it does not already exist.
        /// Returns null if there is no sheet data.
        /// </summary>
        /// <param name="reference">A reference that identifies the cell.</param>
        public Cell? CreateCell(string reference) => CreateCell(new CellReference(reference));

        /// <summary>
        /// Returns the specified cell, creating it if it does not already exist.
        /// Returns null if there is no sheet data.
        /// <param name="reference">A reference that identifies the cell.</param>
        public Cell? CreateCell(CellReference reference)
        {
            Worksheet? worksheet = (reference.SheetName != null) ?
                GetWorksheet(reference.SheetName) :
                Worksheet;

            SheetData? sheetData = worksheet?.GetFirstChild<SheetData>();
            if (sheetData != null)
            {
                // Get or create row
                Row row = CreateRow(sheetData, reference.RowIndex);
                // Get or create cell
                return CreateCell(reference, row);
            }
            return null;
        }

        /// <summary>
        /// Efficiently inserts a collection of contiguous cells into a row. Used by <see cref="TableBuilder"/>.
        /// </summary>
        /// <param name="startReference">Cell reference of starting cell.</param>
        /// <param name="cells">Cells to insert.</param>
        /// <remarks>
        /// The CellReference property will be set for each cell. Any existing CellReference
        /// information is ignored and discarded.
        /// </remarks>
        internal void InsertRowCells(CellReference startReference, IEnumerable<Cell?> cells)
        {
            Worksheet? worksheet = (startReference.SheetName != null) ?
                GetWorksheet(startReference.SheetName) :
                Worksheet;

            SheetData? sheetData = worksheet?.GetFirstChild<SheetData>();
            if (sheetData != null)
            {
                CellReference reference = new(startReference);
                Row row = CreateRow(sheetData, reference.RowIndex);

                // Find initial insert position
                Cell? insertPosition = null;
                uint insertPositionColumnIndex = uint.MaxValue;
                foreach (Cell cell in row)
                {
                    uint columnIndex = CellReference.CellReferenceToColumnIndex(cell.CellReference);
                    if (columnIndex >= reference.ColumnIndex)
                    {
                        insertPosition = cell;
                        insertPositionColumnIndex = columnIndex;
                        break;
                    }
                }

                // Add cells
                foreach (Cell? cell in cells)
                {
                    // Check if we need to replace an existing cell
                    if (reference.ColumnIndex == insertPositionColumnIndex)
                    {
                        Cell? nextPosition = insertPosition!.NextSibling() as Cell;
                        row.RemoveChild(insertPosition);
                        insertPosition = nextPosition;
                        insertPositionColumnIndex = (insertPosition != null) ?
                            CellReference.CellReferenceToColumnIndex(insertPosition.CellReference) :
                            uint.MaxValue;
                    }
                    // Insert new cell
                    if (cell != null)
                    {
                        cell.CellReference = reference.ShortReference;
                        row.InsertBefore(cell, insertPosition);
                    }
                    reference.ColumnIndex++;
                }
            }
        }

        /// <summary>
        /// Creates the specified row or returns the existing row if it already exists.
        /// </summary>
        /// <param name="sheetData">Sheet data where the row should be created.</param>
        /// <param name="rowIndex">Index of the row to create.</param>
        protected Row CreateRow(SheetData sheetData, uint rowIndex)
        {
            Debug.Assert(sheetData != null);

            Row? insertBeforeRow = null;
            foreach (Row row in sheetData.Elements<Row>())
            {
                uint index = row.RowIndex ?? CellReference.DefaultRowIndex;
                if (index == rowIndex)
                    return row;
                if (index > rowIndex)
                {
                    insertBeforeRow = row;
                    break;
                }
            }
            return sheetData.InsertBefore(new Row() { RowIndex = rowIndex }, insertBeforeRow);
        }

        /// <summary>
        /// Creates a cell on the given row or returns the existing cell if it already exists.
        /// </summary>
        /// <param name="reference">Cell reference of the cell to create.</param>
        /// <param name="row">Row to create the cell on.</param>
        protected Cell CreateCell(CellReference reference, Row row)
        {
            Debug.Assert(reference != null);
            Debug.Assert(row != null);

            Cell? insertBeforeCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                uint columnIndex = CellReference.CellReferenceToColumnIndex(cell.CellReference);
                if (columnIndex == reference.ColumnIndex)
                    return cell;
                if (columnIndex > reference.ColumnIndex)
                {
                    insertBeforeCell = cell;
                    break;
                }
            }
            return row.InsertBefore(new Cell() { CellReference = reference.ShortReference }, insertBeforeCell);
        }

        /// <summary>
        /// Deletes the specified cell.
        /// </summary>
        /// <param name="reference">A reference to the cell to delete.</param>
        public void DeleteCell(string reference) => DeleteCell(new CellReference(reference));

        /// <summary>
        /// Deletes the specified cell.
        /// </summary>
        /// <param name="reference">A reference to the cell to delete.</param>
        public void DeleteCell(CellReference reference)
        {
            Debug.Assert(reference != null);

            Cell? cell = FindCell(reference);
            if (cell != null)
                cell.Remove();
        }

        #endregion

        #region Get cell values

        public string? GetCellText(string reference) => GetCellText(FindCell(reference));

        public string? GetCellText(CellReference reference) => GetCellText(FindCell(reference));

        public string? GetCellText(Cell? cell)
        {
            if (cell != null)
            {
                if (cell.CellFormula != null)
                    return $"={cell.CellFormula.InnerText}";

                string value = cell.InnerText;
                if (value != null && cell.DataType != null)
                {
                    switch (cell.DataType.Value)
                    {
                        case CellValues.SharedString:
                            if (int.TryParse(cell.InnerText, out int id))
                                return GetSharedString(id);
                            return value;
                        case CellValues.Boolean:
                            return (value == "0") ? "FALSE" : "TRUE";
                        default:
                            return value;
                    }
                }
            }
            return null;
        }

        public bool? GetCellBoolean(string reference) => GetCellBoolean(FindCell(reference));

        public bool? GetCellBoolean(CellReference reference) => GetCellBoolean(FindCell(reference));

        public bool? GetCellBoolean(Cell? cell)
        {
            if (cell != null && cell.CellValue != null)
                if (cell.CellValue.TryGetBoolean(out bool value))
                    return value;
            return null;
        }

        public int? GetCellInteger(string reference) => GetCellInteger(FindCell(reference));

        public int? GetCellInteger(CellReference reference) => GetCellInteger(FindCell(reference));

        public int? GetCellInteger(Cell? cell)
        {
            if (cell != null && cell.CellValue != null)
                if (cell.CellValue.TryGetInt(out int value))
                    return value;
            return null;
        }

        public double? GetCellDouble(string reference) => GetCellDouble(FindCell(reference));

        public double? GetCellDouble(CellReference reference) => GetCellDouble(FindCell(reference));

        public double? GetCellDouble(Cell? cell)
        {
            if (cell != null && cell.CellValue != null)
                if (cell.CellValue.TryGetDouble(out double value))
                    return value;
            return null;
        }

        public decimal? GetCellDecimal(string reference) => GetCellDecimal(FindCell(reference));

        public decimal? GetCellDecimal(CellReference reference) => GetCellDecimal(FindCell(reference));

        public decimal? GetCellDecimal(Cell? cell)
        {
            if (cell != null && cell.CellValue != null)
                if (cell.CellValue.TryGetDecimal(out decimal value))
                    return value;
            return null;
        }

        public DateTime? GetCellDateTime(string reference) => GetCellDateTime(FindCell(reference));

        public DateTime? GetCellDateTime(CellReference reference) => GetCellDateTime(FindCell(reference));

        public DateTime? GetCellDateTime(Cell? cell)
        {
            if (cell != null && cell.CellValue != null)
                if (cell.CellValue.TryGetDateTime(out DateTime value))
                    return value;
            return null;
        }

        public string? GetCellFormula(string reference) => GetCellFormula(FindCell(reference));

        public string? GetCellFormula(CellReference reference) => GetCellFormula(FindCell(reference));

        public string? GetCellFormula(Cell? cell)
        {
            if (cell != null && cell.CellFormula != null)
                return cell.CellFormula.InnerText;
            return null;
        }

        #endregion

        #region Set cell values

        // String

        public void SetCell(string reference, string value, uint? styleId = null) =>
            SetCell(CreateCell(new CellReference(reference)), value, styleId);

        public void SetCell(CellReference reference, string value, uint? styleId = null) =>
            SetCell(CreateCell(reference), value, styleId);

        public void SetCell(Cell? cell, CellValue<string> value) =>
            SetCell(cell, value.Value, value.StyleId);

        public void SetCell(Cell? cell, string value, uint? styleId = null)
        {
            if (cell != null && value != null)
            {
                int id = AddSharedString(value);
                cell.CellValue = new CellValue(id);
                cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                cell.StyleIndex = styleId ?? CellStyles[StandardCellStyle.General];
            }
        }

        // Integer

        public void SetCell(string reference, int value, uint? styleId = null) =>
            SetCell(CreateCell(new CellReference(reference)), value, styleId);

        public void SetCell(CellReference reference, int value, uint? styleId = null) =>
            SetCell(CreateCell(reference), value, styleId);

        public void SetCell(Cell? cell, CellValue<int> value) =>
            SetCell(cell, value.Value, value.StyleId);

        public void SetCell(Cell? cell, int value, uint? styleId = null)
        {
            if (cell != null)
            {
                cell.CellValue = new CellValue(value);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = styleId ?? CellStyles[StandardCellStyle.Integer];
            }
        }

        // Double

        public void SetCell(string reference, double value, uint? styleId = null) =>
            SetCell(CreateCell(new CellReference(reference)), value, styleId);

        public void SetCell(CellReference reference, double value, uint? styleId = null) =>
            SetCell(CreateCell(reference), value, styleId);

        public void SetCell(Cell? cell, CellValue<double> value) =>
            SetCell(cell, value.Value, value.StyleId);

        public void SetCell(Cell? cell, double value, uint? styleId = null)
        {
            if (cell != null)
            {
                cell.CellValue = new CellValue(value);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = styleId ?? CellStyles[StandardCellStyle.Float];
            }
        }

        // Decimal

        public void SetCell(string reference, decimal value, uint? styleId = null) =>
            SetCell(CreateCell(new CellReference(reference)), value, styleId);

        public void SetCell(CellReference reference, decimal value, uint? styleId = null) =>
            SetCell(CreateCell(reference), value, styleId);

        public void SetCell(Cell? cell, CellValue<decimal> value) =>
            SetCell(cell, value.Value, value.StyleId);

        public void SetCell(Cell? cell, decimal value, uint? styleId = null)
        {
            if (cell != null)
            {
                cell.CellValue = new CellValue(value);
                cell.DataType = new EnumValue<CellValues>(CellValues.Number);
                cell.StyleIndex = styleId ?? CellStyles[StandardCellStyle.Currency];
            }
        }

        // DateTime

        public void SetCell(string reference, DateTime value, uint? styleId = null) =>
            SetCell(CreateCell(new CellReference(reference)), value, styleId);

        public void SetCell(CellReference reference, DateTime value, uint? styleId = null) =>
            SetCell(CreateCell(reference), value, styleId);

        public void SetCell(Cell? cell, CellValue<DateTime> value) =>
            SetCell(cell, value.Value, value.StyleId);

        public void SetCell(Cell? cell, DateTime value, uint? styleId = null)
        {
            if (cell != null)
            {
                cell.CellValue = new CellValue(value);
                cell.DataType = new EnumValue<CellValues>(CellValues.Date);
                cell.StyleIndex = styleId ?? CellStyles[StandardCellStyle.DateTime];
            }
        }

        // CellFormula

        public void SetCell(string reference, CellFormula value, uint? styleId = null, CellValues? dataType = null) =>
            SetCell(CreateCell(new CellReference(reference)), value, styleId, dataType);

        public void SetCell(CellReference reference, CellFormula value, uint? styleId = null, CellValues? dataType = null) =>
            SetCell(CreateCell(reference), value, styleId, dataType);

        public void SetCell(Cell? cell, CellValue<CellFormula> value) =>
            SetCell(cell, value.Value, value.StyleId, value.DataType);

        public void SetCell(Cell? cell, CellFormula value, uint? styleId = null, CellValues? dataType = null)
        {
            if (cell != null)
            {
                cell.CellFormula = value;
                cell.DataType = new EnumValue<CellValues>(dataType ?? CellValues.Number);
                cell.StyleIndex = styleId ?? CellStyles[StandardCellStyle.Integer];
            }
        }

        #endregion

        #region Find cells

        /// <summary>
        /// Returns the specified cell, or null if the cell doesn't exist.
        /// </summary>
        /// <param name="reference">A reference to the cell to find.</param>
        public Cell? FindCell(string reference) => FindCell(new CellReference(reference));

        /// <summary>
        /// Returns the specified cell, or null if the cell doesn't exist.
        /// </summary>
        /// <param name="reference">A reference to the cell to find.</param>
        public Cell? FindCell(CellReference reference)
        {
            Worksheet? worksheet = (reference.SheetName != null) ?
                GetWorksheet(reference.SheetName) :
                Worksheet;

            SheetData? sheetData = worksheet?.GetFirstChild<SheetData>();
            if (sheetData != null)
            {
                Row? row = sheetData.Elements<Row>().Where(r => r.RowIndex?.Value == reference.RowIndex).FirstOrDefault();
                if (row != null)
                    return row.Elements<Cell>().Where(c => c.CellReference == reference.ShortReference).FirstOrDefault();
            }
            return null;
        }

        /// <summary>
        /// Returns the cell with the specified name, or null if the cell wasn't found.
        /// </summary>
        /// <param name="name">Name of the cell to find.</param>
        public Cell? FindCellByName(string name)
        {
            CellReference? reference = GetCellReferenceFromName(name);
            return (reference != null) ?
                FindCell(reference) :
                null;
        }

        /// <summary>
        /// Returns a <see cref="CellReference"/> for the named cell, or null if no cell
        /// was found with the specified name.
        /// </summary>
        /// <param name="name">Name of the cell to find.</param>
        public CellReference? GetCellReferenceFromName(string name)
        {
            DefinedNames? definedNames = WorkbookPart.Workbook?.DefinedNames;
            if (definedNames != null)
            {
                foreach (DefinedName definedName in definedNames)
                {
                    if (string.Compare(definedName.Name, name, true) == 0)
                        return new CellReference(definedName.InnerText);
                }
            }
            return null;
        }

        #endregion

        #region Columns

        /// <summary>
        /// Sets a column width for the active worksheet.
        /// </summary>
        /// <param name="index">1-based index of column to set.</param>
        /// <param name="width">Column width measured as the number of characters of the maximum
        /// digit width as rendered in the normal style's font. There are 4 pixels of margin padding
        /// (two on each side), plus 1 pixel padding for the gridlines.</param>
        public void SetColumnWidth(uint index, double width) => SetColumnWidth(index, index, width);

        /// <summary>
        /// Sets a range of column widths for the active worksheet.
        /// </summary>
        /// <param name="startIndex">1-based index of starting column to set.</param>
        /// <param name="endIndex">1-based index of ending column to set.</param>
        /// <param name="width">Column width measured as the number of characters of the maximum
        /// digit width as rendered in the normal style's font. There are 4 pixels of margin padding
        /// (two on each side), plus 1 pixel padding for the gridlines.</param>
        /// <remarks>
        /// https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.column?view=openxml-2.8.1
        /// </remarks>
        public void SetColumnWidth(uint startIndex, uint endIndex, double width)
        {
            if (Worksheet != null)
            {
                // Get columns
                Columns columns = Worksheet.GetFirstChild<Columns>() ??
                    Worksheet.InsertAt(new Columns(), 0);

                // Add new column info
                columns.Append(new Column()
                {
                    Min = startIndex,
                    Max = endIndex,
                    Width = width,
                    CustomWidth = true,
                });
            }
        }

        #endregion

        #region Tables

        /// <summary>
        /// Returns the table with the specified name, or null of there are no matching tables.
        /// </summary>
        /// <param name="name">The name of the table to find.</param>
        /// <returns>The table with the specified name.</returns>
        public Table? GetTableByName(string name)
        {
            IEnumerable<TableDefinitionPart>? tableDefinitionParts = Worksheet?.WorksheetPart?.TableDefinitionParts;
            if (tableDefinitionParts != null)
            {
                foreach (TableDefinitionPart tableDefinitionPart in tableDefinitionParts)
                {
                    Table table = tableDefinitionPart.Table;
                    if (string.Compare(table.Name, name, true) == 0)
                        return table;
                }
            }
            return null;
        }

        /// <summary>
        /// Creates an Excel table.
        /// </summary>
        /// <param name="name">The name to give the new table.</param>
        /// <param name="range">The cell range of the new table.</param>
        /// <param name="headers">Table headers, if table has header.</param>
        /// <param name="tableStyle">Optional table style.</param>
        /// <returns>Returns the created table.</returns>
        public Table? CreateTable(string name, CellRange range, IEnumerable<string>? headers = null, ExcelTableStyle? tableStyle = null)
        {
            // Get target worksheet
            Worksheet? worksheet = (range.Start.SheetName != null) ?
                GetWorksheet(range.Start.SheetName) :
                Worksheet;

            // Must have worksheet part
            WorksheetPart? worksheetPart = worksheet?.WorksheetPart;
            if (worksheetPart == null)
                return null;

            // Create table definition part
            TableDefinitionPart tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>();

            // Create table
            uint tableId = (uint)(Document.WorkbookPart?.WorksheetParts?.Sum(x => x.TableDefinitionParts.Count()) ?? 0) + 1;
            Table table = new()
            {
                Id = tableId,
                Name = name,
                DisplayName = name,
                Reference = range.ShortReference,
                TotalsRowShown = false,
                HeaderRowCount = headers != null ? 1U : 0U
            };
            tableDefinitionPart.Table = table;

            // Add table columns
            TableColumns tableColumns = new();
            if (headers != null)
            {
                uint id = 1;
                tableColumns.Append(headers.Select(h => new TableColumn()
                {
                    Id = id++,
                    Name = h
                }));
            }
            else
            {
                tableColumns.Append(Enumerable.Range(1, range.ColumnCount).Select(n => new TableColumn()
                {
                    Id = (uint)n,
                    Name = $"Column{n}"
                }));
            }
            tableColumns.Count = (uint)tableColumns.Count();
            table.Append(tableColumns);

            // Add auto filter to table header
            //AutoFilter autoFilter = new() { Reference = range.ShortReference };
            //table.Append(autoFilter);

            // Add table style info
            if (tableStyle != null && !tableStyle.Value.IsEmpty)
            {
                ExcelTableStyle style = tableStyle.Value;
                table.Append(new TableStyleInfo()
                {
                    Name = style.Value,
                    ShowFirstColumn = style.ShowFirstColumn,
                    ShowLastColumn = style.ShowLastColumn,
                    ShowRowStripes = style.ShowRowStripes,
                });
            }

            // Get table parts
            TableParts? tableParts = worksheetPart.Worksheet.Elements<TableParts>().FirstOrDefault();
            if (tableParts == null)
            {
                tableParts = new();
                worksheetPart.Worksheet.Append(tableParts);
            }

            // Add table part
            TablePart tablePart = new() { Id = worksheetPart.GetIdOfPart(tableDefinitionPart) };
            tableParts.Append(tablePart);
            tableParts.Count = (uint)tableParts.Count();

            return table;
        }

        #endregion

        #region Resources

        /// <summary>
        /// Returns the current stylesheet, creating one if necessary.
        /// </summary>
        /// <returns></returns>
        internal Stylesheet GetStylesheet()
        {
            WorkbookStylesPart workbookStylesPart = WorkbookPart.GetPartsOfType<WorkbookStylesPart>()
                ?.FirstOrDefault() ??
                WorkbookPart.AddNewPart<WorkbookStylesPart>();

            if (workbookStylesPart.Stylesheet == null)
                workbookStylesPart.Stylesheet = new();

            return workbookStylesPart.Stylesheet;
        }

        /// <summary>
        /// Adds a shared string resource and returns the index of the shared string. If an entry already
        /// exists for the given text, the index of the existing resource is returned.
        /// </summary>
        /// <param name="text">The string text to add.</param>
        /// <returns>The index of the shared string.</returns>
        internal int AddSharedString(string text)
        {
            // Create SharedStringTablePart if needed
            SharedStringTablePart shareStringPart = WorkbookPart.GetPartsOfType<SharedStringTablePart>()?.FirstOrDefault() ??
                WorkbookPart.AddNewPart<SharedStringTablePart>();

            // Create SharedStringTable if needed
            if (shareStringPart.SharedStringTable == null)
                shareStringPart.SharedStringTable = new();

            // Return index of matching text
            int index = 0;
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                    return index;
                index++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));

            return index;
        }

        /// <summary>
        /// Returns the shared string resource with the specified index, or <c>null</c> if the index is
        /// not valid.
        /// </summary>
        /// <param name="index">The index of the shared string to return.</param>
        /// <returns>The shared string resource with the specified index.</returns>
        internal string? GetSharedString(int index)
        {
            // Create SharedStringTablePart if needed
            SharedStringTablePart shareStringPart = WorkbookPart.GetPartsOfType<SharedStringTablePart>()?.FirstOrDefault() ??
                WorkbookPart.AddNewPart<SharedStringTablePart>();

            // Create SharedStringTable if needed
            if (shareStringPart.SharedStringTable != null && shareStringPart.SharedStringTable.Count() > index)
                return shareStringPart.SharedStringTable.ElementAt(index).InnerText;

            return null;
        }

        #endregion

        #region IDispose

        private bool Disposed = false;

        public void Dispose()
        {
            if (!Disposed)
            {
                Document.Close();
                Document.Dispose();
                Disposed = true;

                GC.SuppressFinalize(this);
            }
        }

        #endregion

    }
}
