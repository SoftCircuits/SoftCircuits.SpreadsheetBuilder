// Copyright (c) 2021 Jonathan Wood (www.softcircuits.com)
// Licensed under the MIT license.
//
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace SoftCircuits.Spreadsheet
{
    public enum StandardNumberingFormat
    {
        General = 164,
        Float = 165,
        DateTime = 166
    }
    
    public class NumberFormats : Dictionary<StandardNumberingFormat, uint>
    {
        // OpenXML reserves numbering formats 0 - 163.
        private const uint MinimumNumberFormatId = 164;

        private readonly SpreadsheetBuilder Builder;

        public NumberFormats(SpreadsheetBuilder builder, bool registerStandardNumberingFormats)
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
