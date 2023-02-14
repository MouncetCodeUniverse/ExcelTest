using OfficeOpenXml;
using System.IO;
using System.Windows.Forms;

namespace ExcelTest
{
    internal static class ExcelUtilities
    {
        internal static void LoadExcelToDataGridView(string filePath, DataGridView dataGridView1)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var totalRows = worksheet.Dimension.End.Row;
                var totalColumns = worksheet.Dimension.End.Column;

                // Set up columns in the DataGridView
                for (int i = 1; i <= totalColumns; i++)
                {
                    dataGridView1.Columns.Add("col" + i, "Column " + i);
                }

                // Populate rows in the DataGridView
                for (int row = 1; row <= totalRows; row++)
                {
                    var dataGridViewRow = new DataGridViewRow();
                    dataGridViewRow.CreateCells(dataGridView1);

                    for (int col = 1; col <= totalColumns; col++)
                    {
                        dataGridViewRow.Cells[col - 1].Value = worksheet.Cells[row, col].Value?.ToString();
                    }

                    dataGridView1.Rows.Add(dataGridViewRow);
                }
            }
        }

    }
}
