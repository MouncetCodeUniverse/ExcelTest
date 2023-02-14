using OfficeOpenXml;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // create excel
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                //Set some properties of the Excel document
                excelPackage.Workbook.Properties.Author = "mena";
                excelPackage.Workbook.Properties.Title = "document 1";
                excelPackage.Workbook.Properties.Subject = "demo 1";
                excelPackage.Workbook.Properties.Created = DateTime.Now;
                
                //Create the WorkSheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
                //MessageBox.Show(excelPackage.Workbook.Worksheets.Count.ToString());
                //Add some text to cell A1
                worksheet.Cells["A1"].Value = "My first EPPlus spreadsheet!";
                worksheet.Cells["B1"].Value = "My second EPPlus spreadsheet!";
                worksheet.Cells["A2"].Value = "my own test";
                worksheet.Cells["B2"].Value = "my own test";
                worksheet.Cells["A3"].Value = "another test";
                worksheet.Cells["A3"].Value = "another test";
                //You could also use [line, column] notation:
                //worksheet.Cells[1, 2].Value = "This is cell B1!";
                //Save your filez
                FileInfo file = new FileInfo(@"document1.xlsx");
                excelPackage.SaveAs(file);
            }
            //Opening an existing Excel file
            FileInfo fi = new FileInfo(@"document1.xlsx");
            using (ExcelPackage excelPackage = new ExcelPackage(fi))
            {
                //Get a WorkSheet by index.
                ExcelWorksheet firstWorksheet = excelPackage.Workbook.Worksheets[0];

                //Get a WorkSheet by name. If the worksheet doesn't exist, throw an exeption
                //ExcelWorksheet namedWorksheet = excelPackage.Workbook.Worksheets["SomeWorksheet"];

                //If you don't know if a worksheet exists, you could use LINQ,
                //So it doesn't throw an exception, but return null in case it doesn't find it
                //ExcelWorksheet anotherWorksheet =
                    //excelPackage.Workbook.Worksheets.FirstOrDefault(x => x.Name == "SomeWorksheet");

                //Get the content from cells A1 and A2 as string, in two different notations
                string valA1 = firstWorksheet.Cells["A1"].Value.ToString();
                string valA2 = firstWorksheet.Cells[2, 1].Value.ToString();

                //Save your file
                excelPackage.Save();
            }
            // import excel file and convert it to list
              List<string> ExcelPackageToDataTable(string filename, int sheetNumber=0)
            {
                List<string> myList = new List<string>();
                FileInfo file = new FileInfo(filename);
                ExcelPackage excelPackage = new ExcelPackage(file);
                DataTable dt = new DataTable();
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[sheetNumber];

                //check if the worksheet is completely empty
                if (worksheet.Dimension == null)
                {
                    return myList;
                }

                //create a list to hold the column names
                List<string> columnNames = new List<string>();

                //needed to keep track of empty column headers
                int currentColumn = 1;

                //loop all columns in the sheet and add them to the datatable
                foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column])
                {
                    string columnName = cell.Text.Trim();

                    //check if the previous header was empty and add it if it was
                    if (cell.Start.Column != currentColumn)
                    {
                        columnNames.Add("Header_" + currentColumn);
                        dt.Columns.Add("Header_" + currentColumn);
                        currentColumn++;
                    }

                    //add the column name to the list to count the duplicates
                    columnNames.Add(columnName);

                    //count the duplicate column names and make them unique to avoid the exception
                    //A column named 'Name' already belongs to this DataTable
                    int occurrences = columnNames.Count(x => x.Equals(columnName));
                    if (occurrences > 1)
                    {
                        columnName = columnName + "_" + occurrences;
                    }

                    //add the column to the datatable
                    dt.Columns.Add(columnName);

                    currentColumn++;
                }

                //start adding the contents of the excel file to the datatable
                for (int i = 2; i <= worksheet.Dimension.End.Row; i++)
                {
                    var row = worksheet.Cells[i, 1, i, worksheet.Dimension.End.Column];
                    DataRow newRow = dt.NewRow();

                    //loop all cells in the row
                    foreach (var cell in row)
                    {
                        newRow[cell.Start.Column - 1] = cell.Text;
                    }

                    dt.Rows.Add(newRow);
                }
                foreach (DataRow dataRow in dt.Rows)
                    myList.Add(string.Join("|", dataRow.ItemArray.Select(item => item.ToString())));
                return myList;
            }
            var test = ExcelPackageToDataTable(@"document1.xlsx",0);
            MessageBox.Show(test[0]);
        }
    }
}
