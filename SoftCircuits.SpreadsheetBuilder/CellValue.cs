// Copyright (c) 2021 Jonathan Wood (www.softcircuits.com)
// Licensed under the MIT license.
//
using DocumentFormat.OpenXml.Spreadsheet;

namespace SoftCircuits.Spreadsheet
{
    /// <summary>
    /// Specifies a cell value and, optionally, an ID and data type.
    /// </summary>
    public class CellValue<T>
    {
        /// <summary>
        /// The value for this <see cref="CellValue"/>.
        /// </summary>
        public T Value { get; set; }

        /// <summary>
        /// The style ID specified for this <see cref="CellValue"/>.
        /// </summary>
        public uint? StyleId { get; set; }

        /// <summary>
        /// The data type specified for this <see cref="CellValue"/>. Only used for formulas.
        /// </summary>
        public CellValues? DataType { get; set; }

        /// <summary>
        /// Constructs a new <see cref="CellValue"/> instance.
        /// </summary>
        /// <param name="value">The value for this <see cref="CellValue"/>.</param>
        /// <param name="styleId">The style ID for this <see cref="CellValue"/>.</param>
        /// <param name="dataType">The data type for this <see cref="CellValue"/>.
        /// This value is ignored for all value types except <see cref="CellFormula"/>.</param>
        public CellValue(T value, uint? styleId = null, CellValues? dataType = null)
        {
            Value = value;
            StyleId = styleId;
            DataType = dataType;
        }

        /// <summary>
        /// Returns a string that represents this object.
        /// </summary>
        public override string ToString() => Value?.ToString() ?? string.Empty;
    }
}
