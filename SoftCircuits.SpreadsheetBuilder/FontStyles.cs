// Copyright (c) 2021 Jonathan Wood (www.softcircuits.com)
// Licensed under the MIT license.
//
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;
using System.Linq;

namespace SoftCircuits.Spreadsheet
{
    public enum StandardFontStyle
    {
        General,
        Bold,
        Header,
        Subheader,
    }

    public class FontStyles : Dictionary<StandardFontStyle, uint>
    {
        private readonly SpreadsheetBuilder Builder;

        public FontStyles(SpreadsheetBuilder builder, bool registerStandardFontStyles)
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
            Fonts fonts = stylesheet.Fonts ?? stylesheet.AppendChild(new Fonts() { KnownFonts = BooleanValue.FromBoolean(true) });
            fonts.Append(font);
            fonts.Count = (uint)fonts.Count();
            return fonts.Count - 1;
        }
    }
}
