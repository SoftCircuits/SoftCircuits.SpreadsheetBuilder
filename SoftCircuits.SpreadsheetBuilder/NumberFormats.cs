// Copyright (c) 2022 Jonathan Wood (www.softcircuits.com)
// Licensed under the MIT license.
//
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace SoftCircuits.Spreadsheet
{
    /// <summary>
    /// Standard Excel numbering formats
    /// </summary>
    public enum StandardNumberingFormat
    {
        General = 164,
        Float = 165,
        DateTime = 166
    }

    /// <summary>
    /// Class to manage registered numbering formats.
    /// </summary>
    public class NumberFormats : Dictionary<StandardNumberingFormat, uint>
    {
        // OpenXML reserves numbering formats 0 - 163.
        private const uint MinimumNumberFormatId = 164;

        private readonly SpreadsheetBuilder Builder;

        /// <summary>
        /// Constructs a new <see cref="NumberFormats"/> instance.
        /// </summary>
        /// <param name="builder">The <see cref="SpreadsheetBuilder"/>.</param>
        /// <param name="registerStandardNumberingFormats">If true, the standard numbering formats
        /// are automatically registered.</param>
        public NumberFormats(SpreadsheetBuilder builder, bool registerStandardNumberingFormats = true)
        {
            Builder = builder;

            // Add standard numbering formats
            if (registerStandardNumberingFormats)
            {
                Add(StandardNumberingFormat.General, Register(new()
                {
                    NumberFormatId = 0, // Set by RegisterFormat
                    FormatCode = "0"
                }));

                Add(StandardNumberingFormat.Float, Register(new()
                {
                    NumberFormatId = 0, // Set by RegisterFormat
                    FormatCode = "#,##0.###"
                }));

                Add(StandardNumberingFormat.DateTime, Register(new()
                {
                    NumberFormatId = 0, // Set by RegisterFormat
                    FormatCode = "m/d/yyyy h:mm AM/PM"
                }));
            }
        }

        /// <summary>
        /// Registers a new <see cref="NumberingFormat"/> and returns its ID.
        /// </summary>
        /// <param name="numberingFormat">The <see cref="NumberingFormat"/> to register.</param>
        /// <returns>The new <see cref="NumberingFormat"/> ID.</returns>
        /// <remarks>
        /// Sets the <c>NumberFormatId</c> property of <paramref name="numberingFormat"/>.
        /// </remarks>
        public uint Register(NumberingFormat numberingFormat)
        {
            Stylesheet stylesheet = Builder.GetStylesheet();
            NumberingFormats numberingFormats = stylesheet.NumberingFormats ?? stylesheet.AppendChild(new NumberingFormats());
            uint numberFormatId = numberingFormats.Elements<NumberingFormat>()
                .Select(f => f.NumberFormatId?.Value ?? 0)
                .DefaultIfEmpty(0U)
                .Max() + 1U;
            if (numberFormatId < MinimumNumberFormatId)
                numberFormatId = MinimumNumberFormatId;
            numberingFormat.NumberFormatId = numberFormatId;
            numberingFormats.Append(numberingFormat);
            numberingFormats.Count = (uint)numberingFormats.Count();
            return numberFormatId;
        }
    }
}
