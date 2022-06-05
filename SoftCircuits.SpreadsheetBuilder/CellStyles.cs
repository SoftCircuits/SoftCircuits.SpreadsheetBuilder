// Copyright (c) 2022 Jonathan Wood (www.softcircuits.com)
// Licensed under the MIT license.
//
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace SoftCircuits.Spreadsheet
{
    /// <summary>
    /// Standard Excel cell styles.
    /// </summary>
    public enum StandardCellStyle
    {
        General,
        Integer,
        Float,
        Currency,
        DateTime,
        Date,
        Time
    }

    /// <summary>
    /// Class to manage registered cell styles.
    /// </summary>
    public class CellStyles : Dictionary<StandardCellStyle, uint>
    {
        private readonly SpreadsheetBuilder Builder;

        /// <summary>
        /// Constructs a new <see cref="CellStyles"/> instance.
        /// </summary>
        /// <param name="builder">The <see cref="SpreadsheetBuilder"/>.</param>
        /// <param name="registerStandardStyles">If true, the standard cell styles are
        /// automatically registered.</param>
        public CellStyles(SpreadsheetBuilder builder, bool registerStandardStyles = true)
        {
            Builder = builder;

            // Add standard cell styles
            if (registerStandardStyles)
            {
                Add(StandardCellStyle.General, Register(new()
                {
                    NumberFormatId = 0,
                    FontId = 0,
                    BorderId = 0,
                    FillId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                }));

                Add(StandardCellStyle.Integer, Register(new()
                {
                    NumberFormatId = (uint)ExcelNumberFormats.Integer,
                    FontId = 0,
                    BorderId = 0,
                    FillId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                }));

                Add(StandardCellStyle.Float, Register(new()
                {
                    NumberFormatId = Builder.NumberFormats[StandardNumberingFormat.Float],
                    FontId = 0,
                    BorderId = 0,
                    FillId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                }));

                Add(StandardCellStyle.Currency, Register(new()
                {
                    NumberFormatId = (uint)ExcelNumberFormats.Currency,
                    FontId = 0,
                    BorderId = 0,
                    FillId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                }));

                Add(StandardCellStyle.DateTime, Register(new()
                {
                    NumberFormatId = Builder.NumberFormats[StandardNumberingFormat.DateTime],
                    FontId = 0,
                    BorderId = 0,
                    FillId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                }));

                Add(StandardCellStyle.Date, Register(new()
                {
                    NumberFormatId = (uint)ExcelNumberFormats.ShortDate,
                    FontId = 0,
                    BorderId = 0,
                    FillId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                }));

                Add(StandardCellStyle.Time, Register(new()
                {
                    NumberFormatId = (uint)ExcelNumberFormats.ShortTime,
                    FontId = 0,
                    BorderId = 0,
                    FillId = 0,
                    ApplyNumberFormat = BooleanValue.FromBoolean(true)
                }));
            }
        }

        /// <summary>
        /// Registers a new <see cref="CellFormat"/> and returns its ID.
        /// </summary>
        /// <param name="format">The <see cref="CellFormat"/> to register.</param>
        /// <returns>The new <see cref="CellFormat"/> ID.</returns>
        public uint Register(CellFormat format)
        {
            Stylesheet stylesheet = Builder.GetStylesheet();
            CellFormats cellFormats = stylesheet.CellFormats ?? stylesheet.AppendChild(new CellFormats());
            cellFormats.Append(format);
            cellFormats.Count = (uint)cellFormats.Count();
            return cellFormats.Count - 1;
        }
    }
}
