// Copyright (c) 2021 Jonathan Wood (www.softcircuits.com)
// Licensed under the MIT license.
//
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace SoftCircuits.Spreadsheet
{
    /// <summary>
    /// Standard Excel fill styles.
    /// </summary>
    public enum StandardFillStyle
    {
        General,
        Gray125
    }

    /// <summary>
    /// Class to manage registered fill styles.
    /// </summary>
    public class FillStyles : Dictionary<StandardFillStyle, uint>
    {
        private readonly SpreadsheetBuilder Builder;

        /// <summary>
        /// Constructs a new <see cref="FillStyles"/> instance.
        /// </summary>
        /// <param name="builder">The <see cref="SpreadsheetBuilder"/>.</param>
        /// <param name="registerStandardFillStyles">If true, the standard fill styles
        /// are automatically registered.</param>
        public FillStyles(SpreadsheetBuilder builder, bool registerStandardFillStyles = true)
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
