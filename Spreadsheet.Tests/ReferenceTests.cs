using Microsoft.VisualStudio.TestTools.UnitTesting;
using SoftCircuits.Spreadsheet;

namespace Spreadsheet.Tests
{
    [TestClass]
    public class ReferenceTests
    {
        [TestMethod]
        public void TestContructors()
        {
            // CellReference()
            CellReference reference = new();
            Assert.AreEqual(reference.ColumnIndex, 1U);
            Assert.AreEqual(reference.RowIndex, 1U);

            // CellReference(uint columnIndex, uint rowIndex)
            foreach (ReferenceData test in ReferenceData.Data)
            {
                reference = new(test.ColumnIndex, test.RowIndex);
                Assert.AreEqual(test.ColumnIndex, reference.ColumnIndex);
                Assert.AreEqual(test.RowIndex, reference.RowIndex);
            }

            // CellReference(string columnName, uint rowIndex)
            foreach (ReferenceData test in ReferenceData.Data)
            {
                reference = new(CellReference.ColumnIndexToName(test.ColumnIndex), test.RowIndex);
                Assert.AreEqual(test.ColumnIndex, reference.ColumnIndex);
                Assert.AreEqual(test.RowIndex, reference.RowIndex);
            }

            // CellReference(string? sheetName, uint columnIndex, uint rowIndex)
            foreach (ReferenceData test in ReferenceData.Data)
            {
                reference = new(test.Sheet, test.ColumnIndex, test.RowIndex);
                Assert.AreEqual(test.Sheet, reference.SheetName);
                Assert.AreEqual(test.ColumnIndex, reference.ColumnIndex);
                Assert.AreEqual(test.RowIndex, reference.RowIndex);
            }

            // CellReference(string? sheetName, uint columnIndex, uint rowIndex, bool fixedColumn, bool fixedRow)
            foreach (ReferenceData test in ReferenceData.Data)
            {
                reference = new(test.Sheet, test.ColumnIndex, test.RowIndex, test.FixedColumn, test.FixedRow);
                Assert.AreEqual(test.Sheet, reference.SheetName);
                Assert.AreEqual(test.ColumnIndex, reference.ColumnIndex);
                Assert.AreEqual(test.RowIndex, reference.RowIndex);
                Assert.AreEqual(test.FixedColumn, reference.FixedColumn);
                Assert.AreEqual(test.FixedRow, reference.FixedRow);
            }

            // CellReference(string? sheetName, string columnName, uint rowIndex)
            foreach (ReferenceData test in ReferenceData.Data)
            {
                reference = new(test.Sheet, CellReference.ColumnIndexToName(test.ColumnIndex), test.RowIndex);
                Assert.AreEqual(test.Sheet, reference.SheetName);
                Assert.AreEqual(test.ColumnIndex, reference.ColumnIndex);
                Assert.AreEqual(test.RowIndex, reference.RowIndex);
            }

            // CellReference(string? sheetName, string columnName, uint rowIndex, bool fixedColumn, bool fixedRow)
            foreach (ReferenceData test in ReferenceData.Data)
            {
                reference = new(test.Sheet, CellReference.ColumnIndexToName(test.ColumnIndex), test.RowIndex, test.FixedColumn, test.FixedRow);
                Assert.AreEqual(test.Sheet, reference.SheetName);
                Assert.AreEqual(test.ColumnIndex, reference.ColumnIndex);
                Assert.AreEqual(test.RowIndex, reference.RowIndex);
                Assert.AreEqual(test.FixedColumn, reference.FixedColumn);
                Assert.AreEqual(test.FixedRow, reference.FixedRow);
            }

            // CellReference(CellReference reference)
            foreach (ReferenceData test in ReferenceData.Data)
            {
                reference = new(new CellReference(test.Name));
                Assert.AreEqual(test.Sheet, reference.SheetName);
                Assert.AreEqual(test.ColumnIndex, reference.ColumnIndex);
                Assert.AreEqual(test.RowIndex, reference.RowIndex);
                Assert.AreEqual(test.FixedColumn, reference.FixedColumn);
                Assert.AreEqual(test.FixedRow, reference.FixedRow);
            }

            // CellReference(string reference)
            foreach (ReferenceData test in ReferenceData.Data)
            {
                reference = new(test.Name);
                Assert.AreEqual(test.Sheet, reference.SheetName);
                Assert.AreEqual(test.ColumnIndex, reference.ColumnIndex);
                Assert.AreEqual(test.RowIndex, reference.RowIndex);
                Assert.AreEqual(test.FixedColumn, reference.FixedColumn);
                Assert.AreEqual(test.FixedRow, reference.FixedRow);
            }
        }
    }
}
