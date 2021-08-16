using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SoftCircuits.Spreadsheet;

namespace SoftCircuits
{
    /// <summary>
    /// Class to track application standard styles.
    /// </summary>
    public class RailtraxStyles
    {
        public uint General { get; private set; }
        public uint Integer { get; private set; }
        public uint Float { get; private set; }
        public uint Currency { get; private set; }
        public uint DateTime { get; private set; }
        public uint Date { get; private set; }
        public uint Time { get; private set; }
        public uint Bold { get; private set; }
        public uint Header { get; private set; }
        public uint Subheader { get; private set; }
        public uint HeaderRight { get; private set; }
        public uint SubheaderRight { get; private set; }

        public ExcelTableStyle HeaderTableStyle { get; private set; }
        public ExcelTableStyle ItemsTableStyle { get; private set; }

        public RailtraxStyles(SpreadsheetBuilder builder)
        {
            General = builder.CellStyles[StandardCellStyle.General];
            Integer = builder.CellStyles[StandardCellStyle.Integer];
            Float = builder.CellStyles[StandardCellStyle.Float];
            Currency = builder.CellStyles[StandardCellStyle.Currency];
            DateTime = builder.CellStyles[StandardCellStyle.DateTime];
            Date = builder.CellStyles[StandardCellStyle.Date];
            Time = builder.CellStyles[StandardCellStyle.Time];

            Bold = builder.CellStyles.Register(new CellFormat()
            {
                FontId = builder.FontStyles[StandardFontStyle.Bold],
                ApplyFont = BooleanValue.FromBoolean(true),
            });

            Header = builder.CellStyles.Register(new CellFormat()
            {
                FontId = builder.FontStyles[StandardFontStyle.Header],
                ApplyFont = BooleanValue.FromBoolean(true)
            });

            Subheader = builder.CellStyles.Register(new CellFormat()
            {
                FontId = builder.FontStyles[StandardFontStyle.Subheader],
                ApplyFont = BooleanValue.FromBoolean(true)
            });

            HeaderRight = builder.CellStyles.Register(new CellFormat()
            {
                FontId = builder.FontStyles[StandardFontStyle.Header],
                ApplyFont = BooleanValue.FromBoolean(true),
                Alignment = new() { Horizontal = HorizontalAlignmentValues.Right },
                ApplyAlignment = BooleanValue.FromBoolean(true)
            });

            SubheaderRight = builder.CellStyles.Register(new CellFormat()
            {
                FontId = builder.FontStyles[StandardFontStyle.Subheader],
                ApplyFont = BooleanValue.FromBoolean(true),
                Alignment = new() { Horizontal = HorizontalAlignmentValues.Right },
                ApplyAlignment = BooleanValue.FromBoolean(true)
            });

            HeaderTableStyle = ExcelTableStyle.MediumWhite4; // .DarkBlue11;
            ItemsTableStyle = ExcelTableStyle.MediumBlue6;
        }
    }
}
