// Copyright (c) 2021 Jonathan Wood (www.softcircuits.com)
// Licensed under the MIT license.
//
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SoftCircuits.Spreadsheet
{
    /// <summary>
    /// Class to make it easier and more efficient to create table data.
    /// Can also create an actual Excel table.
    /// </summary>
    public class TableBuilder
    {
        private readonly SpreadsheetBuilder Builder;
        private readonly CellReference Start;
        private readonly IEnumerable<string>? Headers;
        private readonly int Columns;

        /// <summary>
        /// The row index of the next row added to this table.
        /// </summary>
        public uint RowIndex { get; private set; }

        /// <summary>
        /// Creates a <see cref="TableBuilder"/> instance. Creates a table with no headers.
        /// </summary>
        /// <param name="builder">A reference to the <see cref="SpreadsheetBuilder"/>.</param>
        /// <param name="reference">A reference to the cell at the top, left of this table.</param>
        /// <param name="columns">The number of table columns.</param>
        public TableBuilder(SpreadsheetBuilder builder, string reference, int columns)
            : this(builder, new CellReference(reference), columns)
        {
        }

        /// <summary>
        /// Creates a <see cref="TableBuilder"/> instance. Creates a table with no headers.
        /// </summary>
        /// <param name="builder">A reference to the <see cref="SpreadsheetBuilder"/>.</param>
        /// <param name="reference">A reference to the cell at the top, left of this table.</param>
        /// <param name="columns">The number of table columns.</param>
        public TableBuilder(SpreadsheetBuilder builder, CellReference reference, int columns)
        {
            Builder = builder ?? throw new ArgumentNullException(nameof(builder));
            Start = reference ?? throw new ArgumentNullException(nameof(reference));
            Headers = null;
            Columns = columns;
            RowIndex = Start.RowIndex;
        }

        /// <summary>
        /// Creates a <see cref="TableBuilder"/> instance.
        /// </summary>
        /// <param name="builder">A reference to the <see cref="SpreadsheetBuilder"/>.</param>
        /// <param name="reference">A reference to the cell at the top, left of this table.</param>
        /// <param name="headers">The headers for this table. Also specifies the number of columns.</param>
        public TableBuilder(SpreadsheetBuilder builder, string reference, IEnumerable<string> headers)
            : this(builder, new CellReference(reference), headers)
        {
        }

        /// <summary>
        /// Creates a <see cref="TableBuilder"/> instance.
        /// </summary>
        /// <param name="builder">A reference to the <see cref="SpreadsheetBuilder"/>.</param>
        /// <param name="reference">A reference to the cell at the top, left of this table.</param>
        /// <param name="headers">The headers for this table. Also specifies the number of columns.</param>
        public TableBuilder(SpreadsheetBuilder builder, CellReference reference, IEnumerable<string> headers)
        {
            Builder = builder ?? throw new ArgumentNullException(nameof(builder));
            Start = reference ?? throw new ArgumentNullException(nameof(reference));
            Headers = headers ?? throw new ArgumentNullException(nameof(headers));
            Columns = Headers.Count();
            RowIndex = Start.RowIndex;
            AddRow(headers.ToArray());
        }

        /// <summary>
        /// Returns true if this table has headers.
        /// </summary>
        public bool HasHeader => Headers != null;

        /// <summary>
        /// Adds one row of data to thia table.
        /// </summary>
        /// <remarks>
        /// The number of items in each row does not affect the width of the
        /// defined table.
        /// </remarks>
        public void AddRow(params object[] args)
        {
            List<Cell?> cells = new();
            Cell? cell;

            foreach (object arg in args)
            {
                if (arg != null)
                {
                    cell = new();

                    switch (arg)
                    {
                        case string s:
                            Builder.SetCell(cell, s);
                            break;
                        case int i:
                            Builder.SetCell(cell, i);
                            break;
                        case double dbl:
                            Builder.SetCell(cell, dbl);
                            break;
                        case decimal dec:
                            Builder.SetCell(cell, dec);
                            break;
                        case DateTime dt:
                            Builder.SetCell(cell, dt);
                            break;
                        case CellFormula formula:
                            Builder.SetCell(cell, formula);
                            break;
                        case CellValue<string> stringValue:
                            Builder.SetCell(cell, stringValue);
                            break;
                        case CellValue<int> integerValue:
                            Builder.SetCell(cell, integerValue);
                            break;
                        case CellValue<double> doubleValue:
                            Builder.SetCell(cell, doubleValue);
                            break;
                        case CellValue<decimal> decimalValue:
                            Builder.SetCell(cell, decimalValue);
                            break;
                        case CellValue<DateTime> datetimeValue:
                            Builder.SetCell(cell, datetimeValue);
                            break;
                        case CellValue<CellFormula> formulaValue:
                            Builder.SetCell(cell, formulaValue);
                            break;
                        default:
                            Builder.SetCell(cell, arg.ToString() ?? string.Empty);
                            break;
                    }
                }
                else cell = null;
                cells.Add(cell);
            }

            Builder.InsertRowCells(new(Start.ColumnIndex, RowIndex++), cells);
        }

        /// <summary>
        /// Returns a cell reference to the cell at the top, left of the current table.
        /// </summary>
        /// <returns></returns>
        public CellReference GetStartReference() => Start;

        /// <summary>
        /// Returns a cell reference to the cell at the bottom, right of the current table.
        /// </summary>
        /// <returns></returns>
        public CellReference GetEndReference() => new(Start.ColumnIndex + (uint)Columns - 1, RowIndex - 1);

        /// <summary>
        /// Returns the range for the current table.
        /// </summary>
        /// <param name="includeHeader">Include header in range.</param>
        public CellRange GetTableRange(bool includeHeader = true)
        {
            CellRange range = new(GetStartReference(), GetEndReference());
            if (HasHeader && !includeHeader)
            {
                range.Start.RowIndex++;
                if (range.End.RowIndex < range.Start.RowIndex)
                    range.End.RowIndex = range.Start.RowIndex;
            }
            return range;
        }

        /// <summary>
        /// Returns the range of the specified column within the current table.
        /// </summary>
        /// <param name="columnIndex">0-based column index</param>
        /// <returns></returns>
        public CellRange GetColumnRange(uint columnIndex, bool includeHeader = false) =>
            GetTableRange(includeHeader).GetColumnRange(columnIndex);

        /// <summary>
        /// Returns the range of the specified row within the current table.
        /// </summary>
        /// <param name="rowIndex">0-based row index</param>
        /// <returns></returns>
        public CellRange GetRowRange(uint rowIndex, bool includeHeader = false) =>
            GetTableRange(includeHeader).GetRowRange(rowIndex);

        /// <summary>
        /// Create an Excel spreadsheet table that corresponds to the cells written to this table so far.
        /// builder.
        /// </summary>
        public void BuildTable(string name, ExcelTableStyle tableStyle) =>
            Builder.CreateTable(name, GetTableRange(), Headers, tableStyle);
    }
}
