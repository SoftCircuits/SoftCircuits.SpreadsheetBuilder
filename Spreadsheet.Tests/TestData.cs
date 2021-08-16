using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Spreadsheet.Tests
{
    public class ReferenceData
    {
        public string Name { get; set; }
        public string Sheet { get; set; }
        public uint ColumnIndex { get; set; }
        public uint RowIndex { get; set; }
        public bool FixedColumn { get; set; }
        public bool FixedRow { get; set; }

        public ReferenceData(string name, string sheet, uint columnIndex, uint rowIndex, bool fixedColumn = false, bool fixedRow = false)
        {
            Name = name;
            Sheet = sheet;
            ColumnIndex = columnIndex;
            RowIndex = rowIndex;
            FixedColumn = fixedColumn;
            FixedRow = fixedRow;
        }

        public static readonly ReferenceData[] Data = new[]
        {
            new ReferenceData("A1", null, 1, 1),
            new ReferenceData("A2", null, 1, 2),
            new ReferenceData("A100", null, 1, 100),
            new ReferenceData("A1000", null, 1, 1000),

            new ReferenceData("B1", null, 2, 1),
            new ReferenceData("Z1", null, 26, 1),
            new ReferenceData("AA1", null, 27, 1),
            new ReferenceData("AB1", null, 28, 1),
            new ReferenceData("ZZ1", null, 702, 1),
            new ReferenceData("AAA1", null, 703, 1),

            new ReferenceData("Sheet1!A1", "Sheet1", 1, 1),
            new ReferenceData("Sheet1!A2", "Sheet1", 1, 2),
            new ReferenceData("Sheet1!A100", "Sheet1", 1, 100),
            new ReferenceData("Sheet1!A1000", "Sheet1", 1, 1000),

            new ReferenceData("Sheet1!B1", "Sheet1", 2, 1),
            new ReferenceData("Sheet1!Z1", "Sheet1", 26, 1),
            new ReferenceData("Sheet1!AA1", "Sheet1", 27, 1),
            new ReferenceData("Sheet1!AB1", "Sheet1", 28, 1),
            new ReferenceData("Sheet1!ZZ1", "Sheet1", 702, 1),
            new ReferenceData("Sheet1!AAA1", "Sheet1", 703, 1),

            new ReferenceData("Sheet1!$A$1", "Sheet1", 1, 1, true, true),
            new ReferenceData("Sheet1!$A$2", "Sheet1", 1, 2, true, true),
            new ReferenceData("Sheet1!$A$100", "Sheet1", 1, 100, true, true),
            new ReferenceData("Sheet1!$A$1000", "Sheet1", 1, 1000, true, true),

            new ReferenceData("Sheet1!$B$1", "Sheet1", 2, 1, true, true),
            new ReferenceData("Sheet1!$Z$1", "Sheet1", 26, 1, true, true),
            new ReferenceData("Sheet1!$AA$1", "Sheet1", 27, 1, true, true),
            new ReferenceData("Sheet1!$AB$1", "Sheet1", 28, 1, true, true),
            new ReferenceData("Sheet1!$ZZ$1", "Sheet1", 702, 1, true, true),
            new ReferenceData("Sheet1!$AAA$1", "Sheet1", 703, 1, true, true),
        };
    }
}
