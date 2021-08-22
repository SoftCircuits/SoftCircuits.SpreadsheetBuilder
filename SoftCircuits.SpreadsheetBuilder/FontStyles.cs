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
    /// Standard Excel font styles.
    /// </summary>
    public enum StandardFontStyle
    {
        General,
        Bold,
        Header,
        Subheader,
    }

    /// <summary>
    /// Class to manage registered font styles.
    /// </summary>
    public class FontStyles : Dictionary<StandardFontStyle, uint>
    {
        private readonly SpreadsheetBuilder Builder;

        /// <summary>
        /// Constructs a new <see cref="FontStyles"/> instance.
        /// </summary>
        /// <param name="builder">The <see cref="SpreadsheetBuilder"/>.</param>
        /// <param name="registerStandardFontStyles">If true, the standard font styles
        /// are automatically registered.</param>
        public FontStyles(SpreadsheetBuilder builder, bool registerStandardFontStyles = true)
        {
            Builder = builder;

            // Add standard font styles
            if (registerStandardFontStyles)
            {
                Add(StandardFontStyle.General, Register(new()
                {
                    FontSize = new FontSize() { Val = 11 },
                    FontName = new FontName() { Val = "Calibri" },
                    FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 },
                    FontScheme = new FontScheme() { Val = new EnumValue<FontSchemeValues>(FontSchemeValues.Minor) }
                }));

                Add(StandardFontStyle.Bold, Register(new()
                {
                    FontSize = new FontSize() { Val = 11 },
                    FontName = new FontName() { Val = "Calibri" },
                    FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 },
                    FontScheme = new FontScheme() { Val = new EnumValue<FontSchemeValues>(FontSchemeValues.Minor) },
                    Bold = new Bold(),
                }));

                Add(StandardFontStyle.Header, Register(new()
                {
                    FontSize = new FontSize() { Val = 20 },
                    FontName = new FontName() { Val = "Calibri" },
                    FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 },
                    FontScheme = new FontScheme() { Val = new EnumValue<FontSchemeValues>(FontSchemeValues.Minor) },
                    Bold = new Bold(),
                }));

                Add(StandardFontStyle.Subheader, Register(new()
                {
                    FontSize = new FontSize() { Val = 14 },
                    FontName = new FontName() { Val = "Calibri" },
                    FontFamilyNumbering = new FontFamilyNumbering() { Val = 2 },
                    FontScheme = new FontScheme() { Val = new EnumValue<FontSchemeValues>(FontSchemeValues.Minor) },
                    Bold = new Bold(),
                }));
            }
        }

        /// <summary>
        /// Registers a new <see cref="Font"/> and returns its ID.
        /// </summary>
        /// <param name="font">The <see cref="Font"/> to register.</param>
        /// <returns>The new <see cref="Font"/> ID.</returns>
        public uint Register(Font font)
        {
            Stylesheet stylesheet = Builder.GetStylesheet();
            Fonts fonts = stylesheet.Fonts ?? stylesheet.AppendChild(new Fonts());
            fonts.Append(font);
            fonts.Count = (uint)fonts.Count();
            return fonts.Count - 1;
        }
    }
}
