// Copyright (c) 2021 Jonathan Wood (www.softcircuits.com)
// Licensed under the MIT license.
//
using System;
using System.Diagnostics.CodeAnalysis;

namespace SoftCircuits.Spreadsheet
{
    /// <summary>
    /// Class to specify a cell range.
    /// </summary>
    public class CellRange
    {
        /// <summary>
        /// The <see cref="CellReference"/> of the cell at the top, left of this range.
        /// </summary>
        public CellReference Start { get; set; }

        /// <summary>
        /// The <see cref="CellReference"/> of the cell at the bottom, right of this range.
        /// </summary>
        public CellReference End { get; set; }

        /// <summary>
        /// Constructs a new <see cref="CellRange"/> instance.
        /// </summary>
        public CellRange()
        {
            Start = new();
            End = new();
        }

        /// <summary>
        /// Constructs a new <see cref="CellRange"/> instance.
        /// </summary>
        /// <param name="start">A cell reference to the cell at the top, left of this range.</param>
        /// <param name="end">A cell reference to the cell at the bottom, right of this range.</param>
        public CellRange(CellReference start, CellReference end)
        {
            Start = new(start);
            End = new(end);
        }

        /// <summary>
        /// Constructs a new <see cref="CellRange"/> instance.
        /// </summary>
        /// <param name="range">A range with values to used for this range.</param>
        public CellRange(CellRange range)
        {
            Start = new(range.Start);
            End = new(range.End);
        }

        /// <summary>
        /// Constructs a new <see cref="CellRange"/> instance.
        /// </summary>
        /// <param name="range">A range descriptor to be used for this range.</param>
        public CellRange(string range) => FromRange(range);

        /// <summary>
        /// Sets this range equal to the give range descriptor.
        /// </summary>
        /// <param name="range">A range descriptor to be used for this range.</param>
#if !NETSTANDARD2_0
        [MemberNotNull(nameof(Start))]
        [MemberNotNull(nameof(End))]
#endif
        public void FromRange(string range)
        {
            int pos = range.IndexOf(':');
            if (pos >= 0)
            {
#if NETSTANDARD2_0
                Start = new(range.Substring(0, pos));
                End = new(range.Substring(pos + 1));
#else
                Start = new(range[0..pos]);
                End = new(range[(pos + 1)..]);
#endif
            }
            else
            {
                Start = new(range);
                End = new(Start);
            }
        }

        /// <summary>
        /// Returns true if this range has at least one row and one column.
        /// </summary>
        /// <returns></returns>
        public bool IsValid => End.ColumnIndex >= Start.ColumnIndex && End.RowIndex >= Start.RowIndex;

        /// <summary>
        /// Returns true if this range represents a single cell or less.
        /// </summary>
        /// <returns></returns>
        public bool IsEmpty => End.ColumnIndex <= Start.ColumnIndex && End.RowIndex <= Start.RowIndex;

        /// <summary>
        /// Normalize this range so that the end row and column are greater than or equal to the
        /// start row and column.
        /// </summary>
        public void Normalize()
        {
            if (End.ColumnIndex < Start.ColumnIndex)
            {
                uint temp = End.ColumnIndex;
                End.ColumnIndex = Start.ColumnIndex;
                Start.ColumnIndex = temp;
            }
            if (End.RowIndex < Start.RowIndex)
            {
                uint temp = End.RowIndex;
                End.RowIndex = Start.RowIndex;
                Start.RowIndex = temp;
            }
        }

        /// <summary>
        /// Returns the range of the specified column within this range.
        /// </summary>
        /// <param name="columnIndex">0-based column index relative to
        /// the left column of this range.</param>
        /// <returns></returns>
        public CellRange GetColumnRange(uint columnIndex)
        {
            CellRange range = new(this);
            range.Start.ColumnIndex += columnIndex;
            range.End.ColumnIndex = range.Start.ColumnIndex;
            return range;
        }

        /// <summary>
        /// Returns the range of the specified row within this range.
        /// </summary>
        /// <param name="rowIndex">0-based row index relative to the
        /// top row of this range.</param>
        /// <returns></returns>
        public CellRange GetRowRange(uint rowIndex)
        {
            CellRange range = new(this);
            range.Start.RowIndex += rowIndex;
            range.End.RowIndex = range.Start.RowIndex;
            return range;
        }

        /// <summary>
        /// Returns the index of the top-most row in this range.
        /// </summary>
        public uint TopRowIndex => Start.RowIndex;

        /// <summary>
        /// Returns the index of the bottom-most row in this range.
        /// </summary>
        public uint BottomRowIndex => End.RowIndex;

        /// <summary>
        /// Returns the index of the left-most column in this range.
        /// </summary>
        public uint LeftColumnIndex => Start.ColumnIndex;

        /// <summary>
        /// Returns the index of the right-most column in this range.
        /// </summary>
        public uint RightColumnIndex => End.ColumnIndex;

        /// <summary>
        /// Returns the name of the left-most column in this range.
        /// </summary>
        public string LeftColumnName => CellReference.ColumnIndexToName(Start.ColumnIndex);

        /// <summary>
        /// Returns the name of the right-most column in this range.
        /// </summary>
        public string RightColumnName => CellReference.ColumnIndexToName(End.ColumnIndex);

        /// <summary>
        /// Gets or sets the number of columns in this range.
        /// </summary>
        public int ColumnCount
        {
            get => (int)(End.ColumnIndex - Start.ColumnIndex + 1);
            set => End.ColumnIndex = Start.ColumnIndex + (uint)(Math.Max(value - 1, 0));
        }

        /// <summary>
        /// Gets or sets the number of rows in this range.
        /// </summary>
        public int RowCount
        {
            get => (int)(End.RowIndex - Start.RowIndex + 1);
            set => End.RowIndex = Start.RowIndex + (uint)(value - 1);
        }

        public override string ToString() => $"{Start}:{End}";

        /// <summary>
        /// Returns a string reference to this range.
        /// </summary>
        public string Reference => ToString();

        /// <summary>
        /// Returns a shortened reference to this range.
        /// </summary>
        public string ShortReference => $"{Start.ShortReference}:{End.ShortReference}";
    }
}
