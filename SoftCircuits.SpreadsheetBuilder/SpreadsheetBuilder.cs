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
        public static SaveValidationExceptions ValidationExceptions { get; set; } = SaveValidationExceptions.None;

        public SpreadsheetDocument Document { get; private set; }
        public WorkbookPart WorkbookPart { get; private set; }
        public Worksheet? Worksheet { get; set; }

        public NumberFormats NumberFormats { get; private set; }
        public FontStyles FontStyles { get; private set; }
        public FillStyles FillStyles { get; private set; }
        public BorderStyles BorderStyles { get; private set; }
        public CellStyles CellStyles { get; private set; }

        #region Construction

        public static SpreadsheetBuilder Create(string path, SpreadsheetDocumentType type = SpreadsheetDocumentType.Workbook) =>
            new(SpreadsheetDocument.Create(path, type), true);

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
            NumberFormats = new(this, true);
            FontStyles = new(this, true);
            FillStyles = new(this, true);
            BorderStyles = new(this, true);
            CellStyles = new(this, true);
        }

        #endregion

        #region Saving

        public void Save()
        {
#if DEBUG
            if (ValidationExceptions == SaveValidationExceptions.Always || ValidationExceptions == SaveValidationExceptions.DebugOnly)
#else
            if (ValidationExceptions == SaveValidationExceptions.Always)
#endif
                ThrowExceptionOnValidationErrors(nameof(Save));

            Document.Save();
        }

        public void SaveAs(string path)
        {
#if DEBUG
            if (ValidationExceptions == SaveValidationExceptions.Always || ValidationExceptions == SaveValidationExceptions.DebugOnly)
#else
            if (ValidationExceptions == SaveValidationExceptions.Always)
#endif
                ThrowExceptionOnValidationErrors(nameof(Save));

            Document.SaveAs(path);
        }

        /// <summary>
        /// Performs validation on the current document and returns any found
        /// validation errors. Returns an empty <see cref="IEnumerable{T}"/>
        /// if the current document is valid.
        /// </summary>
        public IEnumerable<ValidationErrorInfo> GetValidationErrors()
        {
            OpenXmlValidator validator = new();
            return validator.Validate(Document);
        }

        private void ThrowExceptionOnValidationErrors(string methodName)
        {
            var errors = GetValidationErrors();
            if (errors.Any())
                throw new Exception($"Validation failed in {methodName}() : {string.Join(", ", errors.Select(e => e.Description))}");
        }

        #endregion

        #region Worksheets

        public Worksheet? GetFirstWorksheet()
        {
            Sheets? sheets = WorkbookPart.Workbook?.GetFirstChild<Sheets>();
            Sheet? sheet = sheets?.Elements<Sheet>().FirstOrDefault();
            if (sheet?.Id?.Value != null)
                return ((WorksheetPart)WorkbookPart.GetPartById(sheet.Id!)).Worksheet;
            return null;
        }

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
        /// Creates a new worksheet
        /// Given a WorkbookPart, inserts a new worksheet.
        /// </summary>
        /// <param name="name">Name for new worksheet.</param>
        /// <returns></returns>
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
        /// Creates the specified cell, or returns the existing cell if it already
        /// exists. Returns null if there is no sheet data.
        /// </summary>
        /// <param name="reference"></param>
        public Cell? CreateCell(string reference) => CreateCell(new CellReference(reference));

        /// <summary>
        /// Creates the specified cell, or returns the existing cell if it already
        /// exists. Returns null if there is no sheet data.
        /// </summary>
        /// <param name="reference"></param>
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
        /// Efficiently inserts a collection of contiguous cells info a row.
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
        /// <param name="reference"></param>
        public void DeleteCell(string reference) => DeleteCell(new CellReference(reference));

        /// <summary>
        /// Deletes the specified cell.
        /// </summary>
        /// <param name="reference"></param>
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
        /// Returns the specified cell, or null if the cell wasn't found.
        /// </summary>
        /// <param name="reference">Cell reference.</param>
        public Cell? FindCell(string reference) => FindCell(new CellReference(reference));

        /// <summary>
        /// Returns the specified cell, or null if the cell wasn't found.
        /// </summary>
        /// <param name="reference">Cell reference.</param>
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
        /// Returns the cell with the specified number, or null if the cell wasn't found.
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
        /// digit width of the numbers 0, 1, 2, …, 9 as rendered in the normal style's font.
        /// There are 4 pixels of margin padding (two on each side), plus 1 pixel padding for
        /// the gridlines.</param>
        public void SetColumnWidth(uint index, double width) => SetColumnWidth(index, index, width);

        /// <summary>
        /// Sets column widths for the active worksheet.
        /// </summary>
        /// <param name="startIndex">1-based index of starting column to set.</param>
        /// <param name="endIndex">1-based index of ending column to set.</param>
        /// <param name="width">Column width measured as the number of characters of the maximum
        /// digit width of the numbers 0, 1, 2, …, 9 as rendered in the normal style's font.
        /// There are 4 pixels of margin padding (two on each side), plus 1 pixel padding for
        /// the gridlines.</param>
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

        public Table? CreateTable(string name, CellRange range, int columns, IEnumerable<string>? headers, ExcelTableStyle tableStyle)
        {
            // Get target worksheet
            Worksheet? worksheet = (range.Start.SheetName != null) ?
                GetWorksheet(range.Start.SheetName) :
                Worksheet;

            // Must have worksheet part
            WorksheetPart? worksheetPart = worksheet?.WorksheetPart;
            if (worksheetPart == null)
                return null;

            // Copy range and force width to match number of columns
            range = new(range);
            if (range.ColumnCount != columns)
                range.ColumnCount = columns;

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
                tableColumns.Append(Enumerable.Range(1, columns).Select(n => new TableColumn()
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
            if (!tableStyle.IsEmpty)
            {
                table.Append(new TableStyleInfo()
                {
                    Name = tableStyle.Value,
                    ShowFirstColumn = tableStyle.ShowFirstColumn,
                    ShowLastColumn = tableStyle.ShowLastColumn,
                    ShowRowStripes = tableStyle.ShowRowStripes,
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
        /// Returns the shared string with the given ID, or <c>null</c> if the ID is not valid.
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        internal string? GetSharedString(int id)
        {
            // Create SharedStringTablePart if needed
            SharedStringTablePart shareStringPart = WorkbookPart.GetPartsOfType<SharedStringTablePart>()?.FirstOrDefault() ??
                WorkbookPart.AddNewPart<SharedStringTablePart>();

            // Create SharedStringTable if needed
            if (shareStringPart.SharedStringTable != null && shareStringPart.SharedStringTable.Count() > id)
                return shareStringPart.SharedStringTable.ElementAt(id).InnerText;

            return null;
        }

        /// <summary>
        /// Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text
        /// and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        /// </summary>
        /// <param name="text"></param>
        /// <returns></returns>
        internal int AddSharedString(string text)
        {
            // Create SharedStringTablePart if needed
            SharedStringTablePart shareStringPart = WorkbookPart.GetPartsOfType<SharedStringTablePart>()?.FirstOrDefault() ??
                WorkbookPart.AddNewPart<SharedStringTablePart>();

            // Create SharedStringTable if needed
            if (shareStringPart.SharedStringTable == null)
                shareStringPart.SharedStringTable = new();

            // Return index of matching text
            int i = 0;
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                    return i;
                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));

            return i;
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
