// Copyright (c) 2021 Jonathan Wood (www.softcircuits.com)
// Licensed under the MIT license.
//
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace SoftCircuits.Spreadsheet
{
    /// <summary>
    /// Standard border styles.
    /// </summary>
    public enum StandardBorderStyle
    {
        General,
    }

    /// <summary>
    /// Dictionary to hold 
    /// </summary>
    public class BorderStyles : Dictionary<StandardBorderStyle, uint>
    {
        private readonly SpreadsheetBuilder Builder;

        public BorderStyles(SpreadsheetBuilder builder, bool registerStandardBorderStyles)
        {
            Builder = builder;

            // Add standard border styles
            if (registerStandardBorderStyles)
            {
                Add(StandardBorderStyle.General, Register(new()
                {
                    LeftBorder = new LeftBorder(),
                    RightBorder = new RightBorder(),
                    TopBorder = new TopBorder(),
                    BottomBorder = new BottomBorder(),
                    DiagonalBorder = new DiagonalBorder()
                }));
            }
        }

        /// <summary>
        /// Registers a new <see cref="Border"/> and returns its ID.
        /// </summary>
        /// <param name="border">The <see cref="Border"/> to register.</param>
        /// <returns>The new <see cref="Border"/> ID.</returns>
        public uint Register(Border border)
        {
            Stylesheet stylesheet = Builder.GetStylesheet();
            Borders borders = stylesheet.Borders ?? stylesheet.AppendChild(new Borders());
            borders.Append(border);
            borders.Count = (uint)borders.Count();
            return borders.Count - 1;
        }
    }
}
