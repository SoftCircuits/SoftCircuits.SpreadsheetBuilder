// Copyright (c) 2021 Jonathan Wood (www.softcircuits.com)
// Licensed under the MIT license.
//
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace SoftCircuits.Spreadsheet
{
    public enum StandardFillStyle
    {
        General,
        Gray125
    }
    
    public class FillStyles : Dictionary<StandardFillStyle, uint>
    {
        private readonly SpreadsheetBuilder Builder;

        public FillStyles(SpreadsheetBuilder builder, bool registerStandardFillStyles)
        {
            Builder = builder;

            // Add standard fill styles
            if (registerStandardFillStyles)
            {
                Add(StandardFillStyle.General, Register(new()
                {
                    PatternFill = new PatternFill() { PatternType = new EnumValue<PatternValues>(PatternValues.None) },
                }));

                Add(StandardFillStyle.Gray125, Register(new()
                {
                    PatternFill = new PatternFill() { PatternType = new EnumValue<PatternValues>(PatternValues.Gray125) },
                }));
            }
        }

        /// <summary>
        /// Registers a new <see cref="Fill"/> and returns its ID.
        /// </summary>
        /// <param name="fill">The <see cref="Fill"/> to register.</param>
        /// <returns>The new <see cref="Fill"/> ID.</returns>
        public uint Register(Fill fill)
        {
            Stylesheet stylesheet = Builder.GetStylesheet();
            Fills fills = stylesheet.Fills ?? stylesheet.AppendChild(new Fills());
            fills.Append(fill);
            fills.Count = (uint)fills.Count();
            return fills.Count - 1;
        }
    }
}
