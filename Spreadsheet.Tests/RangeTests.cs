using Microsoft.VisualStudio.TestTools.UnitTesting;
using SoftCircuits.Spreadsheet;

namespace Spreadsheet.Tests
{
    [TestClass]
    public class RangeTests
    {
        [TestMethod]
        public void Test1()
        {
            ReferenceData[] data = ReferenceData.Data;

            for (int i = 0; i < data.Length; i++)
            {
                for (int j = 0; j < data.Length; j++)
                {
                    string s = $"{data[i].Name}:{data[j].Name}";

                    CellRange range = new(s);

                    Assert.AreEqual(data[i].Sheet, range.Start.SheetName);
                    Assert.AreEqual(data[i].ColumnIndex, range.Start.ColumnIndex);
                    Assert.AreEqual(data[i].RowIndex, range.Start.RowIndex);
                    Assert.AreEqual(data[i].FixedColumn, range.Start.FixedColumn);
                    Assert.AreEqual(data[i].FixedRow, range.Start.FixedRow);

                    Assert.AreEqual(data[j].Sheet, range.End.SheetName);
                    Assert.AreEqual(data[j].ColumnIndex, range.End.ColumnIndex);
                    Assert.AreEqual(data[j].RowIndex, range.End.RowIndex);
                    Assert.AreEqual(data[j].FixedColumn, range.End.FixedColumn);
                    Assert.AreEqual(data[j].FixedRow, range.End.FixedRow);

                    if (!range.IsValid)
                    {
                        range.Normalize();
                        Assert.IsTrue(range.IsValid);
                    }
                }
            }
        }
    }
}
