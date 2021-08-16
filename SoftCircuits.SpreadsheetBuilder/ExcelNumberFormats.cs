// Copyright (c) 2021 Jonathan Wood (www.softcircuits.com)
// Licensed under the MIT license.
//
namespace SoftCircuits.Spreadsheet
{
    public enum ExcelNumberFormats
    {
        /// <summary>
        /// General
        /// </summary>
        General = 0,

        /// <summary>
        /// 0
        /// </summary>
        Number = 1,

        /// <summary>
        /// 0.00
        /// </summary>
        TwoDecimalsNoComma = 2,

        /// <summary>
        /// #,##0
        /// </summary>
        Integer = 3,

        /// <summary>
        /// #,##0.00
        /// </summary>
        TwoDecimals = 4,

        /// <summary>
        /// 0%
        /// </summary>
        Percent = 9,

        /// <summary>
        /// 0.00%
        /// </summary>
        PercentHundredths = 10,

        /// <summary>
        /// 0.00E+00
        /// </summary>
        Exponential = 11,

        /// <summary>
        /// # ?/?
        /// </summary>
        ShortFraction = 12,

        /// <summary>
        /// # ??/??
        /// </summary>
        Fraction = 13,

        /// <summary>
        /// d/m/yyyy
        /// </summary>
        ShortDate = 14,

        /// <summary>
        /// d-mmm-yy
        /// </summary>
        MediumDate = 15,

        /// <summary>
        /// d-mmm
        /// </summary>
        DayMonth = 16,

        /// <summary>
        /// mmm-yy
        /// </summary>
        MonthYear = 17,

        /// <summary>
        /// h:mm tt
        /// </summary>
        ShortTime = 18,

        /// <summary>
        /// h:mm:ss tt
        /// </summary>
        LongTime = 19,

        /// <summary>
        /// H:mm
        /// </summary>
        HourMinutes = 20,

        /// <summary>
        /// H:mm:ss
        /// </summary>
        MediumTime = 21,

        /// <summary>
        /// m/d/yyyy H:mm
        /// </summary>
        DateTime = 22,

        /// <summary>
        /// #,##0 ;(#,##0)
        /// </summary>
        ParenNegativeInteger = 37,

        /// <summary>
        /// #,##0 ;[Red](#,##0)
        /// </summary>
        RedNegativeInteger = 38,

        /// <summary>
        /// #,##0.00;(#,##0.00)
        /// </summary>
        ParenNegativeFloat = 39,

        /// <summary>
        /// #,##0.00;[Red](#,##0.00)
        /// </summary>
        RedNegativeFloat = 40,

        /// <summary>
        /// "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)"
        /// </summary>
        Currency = 44,

        /// <summary>
        /// mm:ss
        /// </summary>
        MinuteSeconds = 45,

        /// <summary>
        /// [h]:mm:ss
        /// </summary>
        ElapsedTimeMinuteSeconds = 46,

        /// <summary>
        /// mmss.0
        /// </summary>
        x = 47,

        /// <summary>
        /// ##0.0E+0
        /// </summary>
        Exponential2 = 48,

        /// <summary>
        /// @
        /// </summary>
        PlaceHolder = 49,
    }
}
