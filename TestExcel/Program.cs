using DocumentFormat.OpenXml.Spreadsheet;
using SoftCircuits.Spreadsheet;
using System;

namespace SoftCircuits
{
    class Program
    {
        static readonly string Filename = @"D:\Users\jwood\Documents\TestExcel.xlsx";

        static void Main(string[] args)
        {
            SpreadsheetBuilder.ValidationExceptions = SaveValidationExceptions.None;

            using SpreadsheetBuilder builder = SpreadsheetBuilder.Create(Filename);
            RailtraxStyles styles = new(builder);

            // Table columns
            string[] bodyHeaders = new string[]
            {
                "Railcar",
                "Equipment Type",
                "Model",
                "Start Date",
                "End Date",
                "Monthly Rate",
                "Total"
            };

            // Create 10 sheets
            for (int i = 1; i <= 10; i++)
            {
                if (i > 1)
                    builder.Worksheet = builder.CreateWorksheet($"Sheet{i}");

                builder.SetColumnWidth(1, (uint)bodyHeaders.Length, 18);

                CellReference reference = new(1, 1);

                builder.SetCell("A1", "Company Name", styles.Header);
                builder.SetCell("G1", "Invoice 1117", styles.HeaderRight);
                builder.SetCell("A2", "Customer Name", styles.Subheader);
                builder.SetCell("G2", "6/1/2021 - 6/30/2021", styles.SubheaderRight);

                // Header table
                TableBuilder table = new(builder, "A4", 2);
                table.AddRow(new CellValue<string>("Total Cars:", styles.Bold), 16);
                table.AddRow(new CellValue<string>("Total:", styles.Bold), 4000m);
                table.BuildTable($"HeaderTable{i}", styles.HeaderTableStyle);

                // Items table
                table = new(builder, "A7", bodyHeaders);
                for (int j = 0; j <= 10; j++)
                {
                    table.AddRow("TILX332106",
                        "Covered Hopper",
                        "3281 Cubi Foot Hopper",
                        new CellValue<DateTime>(new DateTime(2020, 10, 1), styles.Date),
                        null,
                        (decimal)250.00,
                        new CellValue<CellFormula>(new CellFormula($"F{table.RowIndex}"), styles.Currency));
                }

                CellRange railcarRange = table.GetColumnRange(1);
                CellRange totalRange = table.GetColumnRange((uint)(bodyHeaders.Length - 1));
                table.AddRow(new CellFormula($"ROWS({railcarRange})"),
                    null,
                    null,
                    null,
                    null,
                    null,
                    new CellValue<CellFormula>(new($"SUM({totalRange})"), styles.Currency));
                table.BuildTable($"ItemsTable{i}", styles.ItemsTableStyle);
            }
            builder.Save();
        }
    }
}
