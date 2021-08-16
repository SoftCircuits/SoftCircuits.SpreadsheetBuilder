// Copyright (c) 2021 Jonathan Wood (www.softcircuits.com)
// Licensed under the MIT license.
//
using DocumentFormat.OpenXml;
using System;
using System.Text;

namespace SoftCircuits.Spreadsheet
{
    /// <summary>
    /// Class to specify a cell reference.
    /// </summary>
    public class CellReference
    {
        public const string DefaultSheetName = null;
        public const uint DefaultColumnIndex = 1U;
        public const uint DefaultRowIndex = 1U;
        public const string DefaultColumnName = "A";

        public string? SheetName { get; set; }
        public uint ColumnIndex { get; set; }
        public uint RowIndex { get; set; }

        public bool FixedColumn { get; set; }
        public bool FixedRow { get; set; }

        /// <summary>
        /// Constructs a new <see cref="CellReference"/> instance.
        /// </summary>
        public CellReference()
        {
            SheetName = DefaultSheetName;
            ColumnIndex = DefaultColumnIndex;
            RowIndex = DefaultRowIndex;
            FixedColumn = false;
            FixedRow = false;
        }

        /// <summary>
        /// Constructs a new <see cref="CellReference"/> instance.
        /// </summary>
        /// <param name="columnIndex">1-based column index.</param>
        /// <param name="rowIndex">1-based row index.</param>
        public CellReference(uint columnIndex, uint rowIndex)
        {
            SheetName = DefaultSheetName;
            ColumnIndex = columnIndex;
            RowIndex = rowIndex;
            FixedColumn = false;
            FixedRow = false;
        }

        /// <summary>
        /// Constructs a new <see cref="CellReference"/> instance.
        /// </summary>
        /// <param name="columnName">Column name.</param>
        /// <param name="rowIndex">1-based row index.</param>
        public CellReference(string columnName, uint rowIndex)
        {
            SheetName = DefaultSheetName;
            ColumnIndex = ColumnNameToIndex(columnName);
            RowIndex = rowIndex;
            FixedColumn = false;
            FixedRow = false;
        }

        /// <summary>
        /// Constructs a new <see cref="CellReference"/> instance.
        /// </summary>
        /// <param name="sheetName">Worksheet name.</param>
        /// <param name="columnIndex">1-based column index.</param>
        /// <param name="rowIndex">1-based row index.</param>
        public CellReference(string? sheetName, uint columnIndex, uint rowIndex)
        {
            SheetName = sheetName;
            ColumnIndex = columnIndex;
            RowIndex = rowIndex;
            FixedColumn = false;
            FixedRow = false;
        }

        /// <summary>
        /// Constructs a new <see cref="CellReference"/> instance.
        /// </summary>
        /// <param name="sheetName">Worksheet name.</param>
        /// <param name="columnIndex">1-based column index.</param>
        /// <param name="rowIndex">1-based row index.</param>
        /// <param name="fixedColumn">Specifies if <paramref name="columnIndex"/> is fixed.</param>
        /// <param name="fixedRow">Specifies if <paramref name="fixedColumn"/> is fixed.</param>
        public CellReference(string? sheetName, uint columnIndex, uint rowIndex, bool fixedColumn, bool fixedRow)
        {
            SheetName = sheetName;
            ColumnIndex = columnIndex;
            RowIndex = rowIndex;
            FixedColumn = fixedColumn;
            FixedRow = fixedRow;
        }

        /// <summary>
        /// Constructs a new <see cref="CellReference"/> instance.
        /// </summary>
        /// <param name="sheetName">Worksheet name.</param>
        /// <param name="columnName">Column name.</param>
        /// <param name="rowIndex">1-based row index.</param>
        public CellReference(string? sheetName, string columnName, uint rowIndex)
        {
            SheetName = sheetName;
            ColumnIndex = ColumnNameToIndex(columnName);
            RowIndex = rowIndex;
            FixedColumn = false;
            FixedRow = false;
        }

        /// <summary>
        /// Constructs a new <see cref="CellReference"/> instance.
        /// </summary>
        /// <param name="sheetName">Worksheet name.</param>
        /// <param name="columnName">Column name.</param>
        /// <param name="rowIndex">1-based row index.</param>
        /// <param name="fixedColumn">Specifies if <paramref name="columnName"/> is fixed.</param>
        /// <param name="fixedRow">Specifies if <paramref name="fixedColumn"/> is fixed.</param>
        public CellReference(string? sheetName, string columnName, uint rowIndex, bool fixedColumn, bool fixedRow)
        {
            SheetName = sheetName;
            ColumnIndex = ColumnNameToIndex(columnName);
            RowIndex = rowIndex;
            FixedColumn = fixedColumn;
            FixedRow = fixedRow;
        }

        /// <summary>
        /// Constructs a new <see cref="CellReference"/> instance.
        /// </summary>
        /// <param name="reference">A <see cref="CellReference"/> with the values to copy
        /// to this instance.</param>
        public CellReference(CellReference reference)
            : this(reference.SheetName, reference.ColumnIndex, reference.RowIndex)
        {
            FixedColumn = reference.FixedColumn;
            FixedRow = reference.FixedRow;
        }

        /// <summary>
        /// Constructs a new <see cref="CellReference"/> instance.
        /// </summary>
        /// <param name="reference">A cell reference string with the values for this
        /// instance.
        /// </param>
        public CellReference(string reference) => FromReference(reference);

        public void FromReference(string reference)
        {
            int pos, start;

            if (reference == null)
                throw new ArgumentNullException(nameof(reference));

            // Parse sheet name
            pos = reference.IndexOf('!');
            if (pos >= 0)
            {
#if NETSTANDARD2_0
                SheetName = reference.Substring(0, pos);
#else
                SheetName = reference[0..pos];
#endif
                pos++;
            }
            else
            {
                SheetName = DefaultSheetName;
                pos = 0;
            }

            // Parse column name
            if (pos < reference.Length && reference[pos] == '$')
            {
                FixedColumn = true;
                pos++;
            }
            start = pos;
            while (pos < reference.Length && char.IsLetter(reference[pos]))
                pos++;
#if NETSTANDARD2_0
            ColumnIndex = ColumnNameToIndex(reference.Substring(start, pos));
#else
            ColumnIndex = ColumnNameToIndex(reference[start..pos]);
#endif

            // Parse row index
            if (pos < reference.Length && reference[pos] == '$')
            {
                FixedRow = true;
                pos++;
            }
            start = pos;
            while (pos < reference.Length && char.IsDigit(reference[pos]))
                pos++;
#if NETSTANDARD2_0
            RowIndex = uint.TryParse(reference.Substring(start, pos), out uint value) ?
#else
            RowIndex = uint.TryParse(reference[start..pos], out uint value) ?
#endif
                value :
                DefaultRowIndex;
        }

        /// <summary>
        /// Converts the value of this instance to a string.
        /// </summary>
        public override string ToString()
        {
            StringBuilder builder = new();

            if (SheetName != null)
            {
                builder.Append(SheetName);
                builder.Append('!');
            }

            if (FixedColumn)
                builder.Append('$');
            builder.Append(ColumnIndexToName(ColumnIndex));

            if (FixedRow)
                builder.Append('$');
            builder.Append(RowIndex);

            return builder.ToString();
        }

        public string ColumnName => ColumnIndexToName(ColumnIndex);

        public string Reference => ToString();

        public string ShortReference => $"{ColumnIndexToName(ColumnIndex)}{RowIndex}";

#region Column conversions

        private const uint AlphabetLength = 26U;

        /// <summary>
        /// Converts a column index to its corresponding name value.
        /// </summary>
        /// <param name="value">Integer to convert</param>
        public static string ColumnIndexToName(uint columnIndex)
        {
            if (columnIndex > 0)
            {
                StringBuilder builder = new(6);

                do
                {
                    columnIndex--;
                    char c = (char)('A' + columnIndex % AlphabetLength);
                    builder.Insert(0, c);
                    columnIndex /= AlphabetLength;
                } while (columnIndex > 0);

                return builder.ToString();
            }
            return DefaultColumnName;
        }

        /// <summary>
        /// Converts a column name to its corresponding index value.
        /// </summary>
        /// <param name="columnName">Column name to convert.</param>
        public static uint ColumnNameToIndex(string? columnName)
        {
            uint columnIndex = 0;
            int pos = 0;
            char c;

            if (columnName != null && columnName.Length > 0 && char.IsUpper(c = columnName[pos]))
            {
                do
                {
                    int i = c - 'A';
                    columnIndex *= AlphabetLength;
                    columnIndex += (uint)i + 1U;
                    pos++;
                } while (pos < columnName.Length && char.IsUpper(c = columnName[pos]));

                return columnIndex;
            }
            return DefaultColumnIndex;
        }

        /// <summary>
        /// Gets a column index from a cell reference.
        /// </summary>
        /// <param name="cellReference">An OpenXML cell reference of type <see cref="StringValue"/>.</param>
        public static uint CellReferenceToColumnIndex(StringValue? cellReference) => ColumnNameToIndex(cellReference?.Value);

#endregion

        public static implicit operator string(CellReference cell) => cell.ToString();
    }
}
