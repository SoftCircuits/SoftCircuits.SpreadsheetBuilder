// Copyright (c) 2021 Jonathan Wood (www.softcircuits.com)
// Licensed under the MIT license.
//
namespace SoftCircuits.Spreadsheet
{
    public struct ExcelTableStyle
    {
        public static readonly ExcelTableStyle Empty = new(null);

        public static ExcelTableStyle LightWhite1 => new("TableStyleLight1");
        public static ExcelTableStyle LightBlue2 => new("TableStyleLight2");
        public static ExcelTableStyle LightOrange3 => new("TableStyleLight3");
        public static ExcelTableStyle LightWhite4 => new("TableStyleLight4");
        public static ExcelTableStyle LightYellow5 => new("TableStyleLight5");
        public static ExcelTableStyle LightBlue6 => new("TableStyleLight6");
        public static ExcelTableStyle LightGreen7 => new("TableStyleLight7");
        public static ExcelTableStyle LightWhite8 => new("TableStyleLight8");
        public static ExcelTableStyle LightBlue9 => new("TableStyleLight9");
        public static ExcelTableStyle LightOrange10 => new("TableStyleLight10");
        public static ExcelTableStyle LightWhite11 => new("TableStyleLight11");
        public static ExcelTableStyle LightGold12 => new("TableStyleLight12");
        public static ExcelTableStyle LightBlue13 => new("TableStyleLight13");
        public static ExcelTableStyle LightGreen14 => new("TableStyleLight14");
        public static ExcelTableStyle LightWhite15 => new("TableStyleLight15");
        public static ExcelTableStyle LightBlue16 => new("TableStyleLight16");
        public static ExcelTableStyle LightOrange17 => new("TableStyleLight17");
        public static ExcelTableStyle LightWhite18 => new("TableStyleLight18");
        public static ExcelTableStyle LightYellow19 => new("TableStyleLight19");
        public static ExcelTableStyle LightBlue20 => new("TableStyleLight20");
        public static ExcelTableStyle LightGreen21 => new("TableStyleLight21");

        public static ExcelTableStyle MediumWhite1 => new("TableStyleMedium1");
        public static ExcelTableStyle MediumBlue2 => new("TableStyleMedium2");
        public static ExcelTableStyle MediumOrange3 => new("TableStyleMedium3");
        public static ExcelTableStyle MediumWhite4 => new("TableStyleMedium4");   // *
        public static ExcelTableStyle MediumGold5 => new("TableStyleMedium5");
        public static ExcelTableStyle MediumBlue6 => new("TableStyleMedium6");    // *
        public static ExcelTableStyle MediumGreen7 => new("TableStyleMedium7");
        public static ExcelTableStyle MediumWhite8 => new("TableStyleMedium8");
        public static ExcelTableStyle MediumBlue9 => new("TableStyleMedium9");
        public static ExcelTableStyle MediumOrange10 => new("TableStyleMedium10");
        public static ExcelTableStyle MediumWhite11 => new("TableStyleMedium11");
        public static ExcelTableStyle MediumGold12 => new("TableStyleMedium12");
        public static ExcelTableStyle MediumBlue13 => new("TableStyleMedium13");
        public static ExcelTableStyle MediumGreen14 => new("TableStyleMedium14");
        public static ExcelTableStyle MediumWhite15 => new("TableStyleMedium15");
        public static ExcelTableStyle MediumBlue16 => new("TableStyleMedium16");
        public static ExcelTableStyle MediumOrange17 => new("TableStyleMedium17");
        public static ExcelTableStyle MediumWhite18 => new("TableStyleMedium18");
        public static ExcelTableStyle MediumGold19 => new("TableStyleMedium19");
        public static ExcelTableStyle MediumBlue20 => new("TableStyleMedium20");
        public static ExcelTableStyle MediumGreen21 => new("TableStyleMedium21");
        public static ExcelTableStyle MediumGray22 => new("TableStyleMedium22");
        public static ExcelTableStyle MediumBlue23 => new("TableStyleMedium23");
        public static ExcelTableStyle MediumOrange24 => new("TableStyleMedium24");
        public static ExcelTableStyle MediumGray25 => new("TableStyleMedium25");
        public static ExcelTableStyle MediumYellow26 => new("TableStyleMedium26");
        public static ExcelTableStyle MediumBlue27 => new("TableStyleMedium27");
        public static ExcelTableStyle MediumGreen28 => new("TableStyleMedium28");

        public static ExcelTableStyle DarkGray1 => new("TableStyleDark1");
        public static ExcelTableStyle DarkBlue2 => new("TableStyleDark2");
        public static ExcelTableStyle DarkBrown3 => new("TableStyleDark3");
        public static ExcelTableStyle DarkGray4 => new("TableStyleDark4");
        public static ExcelTableStyle DarkYellow5 => new("TableStyleDark5");
        public static ExcelTableStyle DarkBlue6 => new("TableStyleDark6");
        public static ExcelTableStyle DarkGreen7 => new("TableStyleDark7");
        public static ExcelTableStyle DarkGray8 => new("TableStyleDark8");
        public static ExcelTableStyle DarkBlue9 => new("TableStyleDark9");
        public static ExcelTableStyle DarkGray10 => new("TableStyleDark10");
        public static ExcelTableStyle DarkBlue11 => new("TableStyleDark11");

        public bool IsEmpty => string.IsNullOrEmpty(Value);

        public readonly string? Value;
        public readonly bool ShowRowStripes;
        public readonly bool ShowFirstColumn;
        public readonly bool ShowLastColumn;

        internal ExcelTableStyle(string? style, bool showRowStripes = true, bool showFirstColumn = false, bool showLastColumn = false)
        {
            Value = style;
            ShowRowStripes = showRowStripes;
            ShowFirstColumn = showFirstColumn;
            ShowLastColumn = showLastColumn;
        }
    }
}
