using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace MergeExcelsTables
{
    internal class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            string[] filePaths;

            if (args.Length < 2)
            {
                // Launch the file picker dialog to allow the user to select files
                filePaths = LaunchFilePicker();
                if (filePaths == null)
                {
                    Console.WriteLine("Error: No files selected.");
                    return -1;
                }
            }
            else if (!args.All(path => File.Exists(path)))
            {
                return -1;
            }
            else
            {
                filePaths = args;
            }

            Application excelApp = new Application { Visible = false };
            Workbook newWorkbook = excelApp.Workbooks.Add();

            try
            {
                bool firstFile = true;
                int currentRow = 2;
                foreach (string filePath in filePaths)
                {
                    string fileExt = Path.GetExtension(filePath).ToLower();
                    Workbook workbook;

                    if (fileExt == ".xlsx" || fileExt == ".xlsb" || fileExt == "xlsm" || fileExt == "xls")
                    {
                        workbook = excelApp.Workbooks.Open(filePath);
                    }
                    else if (fileExt == ".txt")
                    {
                        workbook = CreateExcelWorkbookFromTextFile2(excelApp, filePath);
                    }
                    else
                    {
                        excelApp.Workbooks.Close();
                        return -1;
                    }

                    if (workbook == null)
                    {
                        excelApp.Workbooks.Close();
                        return -1;
                    }

                    // Check if the first row is the same and if there are any excess columns
                    Worksheet worksheet = workbook.Worksheets[1];
                    Range usedRange = worksheet.UsedRange;
                    Range headerRow = usedRange.Rows[1];
                    int lastColumnNumber = (int)excelApp.WorksheetFunction.CountA(headerRow.Rows[1]);
                    headerRow = usedRange.Range["A1",usedRange.Cells[1,lastColumnNumber]];

                    if (firstFile)
                    {
                        headerRow.Copy(newWorkbook.Worksheets[1].Rows[1]);
                        firstFile = false;
                    }
                    else if (!RowsEqual(headerRow, newWorkbook.Worksheets[1].Rows[1]))
                    {
                        throw new Exception();
                    }

                    if (HasExcessColumns(usedRange, lastColumnNumber))
                    {
                        throw new Exception();
                    }

                    /*// Copy the used range below the first row to the new workbook
                    usedRange.Offset[1, 0].Resize[usedRange.Rows.Count - 1].Copy(newWorkbook.Worksheets[1].Cells[currentRow, 1]);
                    currentRow += usedRange.Rows.Count - 1;*/
                    usedRange.Offset[1, 0].Resize[usedRange.Rows.Count - 1].Copy();
                    newWorkbook.Worksheets[1].Cells[currentRow, 1].PasteSpecial(XlPasteType.xlPasteFormulasAndNumberFormats);
                    currentRow += usedRange.Rows.Count - 1;

                    workbook.Close(false);
                }
                Worksheet finalWorksheet = newWorkbook.Worksheets[1];
                Range finalRange = finalWorksheet.UsedRange;
                // Format the range as a table with headers
                ListObject table = finalWorksheet.ListObjects.AddEx(XlListObjectSourceType.xlSrcRange,
                    finalRange, Type.Missing, XlYesNoGuess.xlYes, Type.Missing);
                table.Name = "MyTable";
                table.TableStyle = "TableStyleLight1";

                string newFilePath = PromptForSaveLocation();
                newWorkbook.Application.DisplayAlerts = false;
                newWorkbook.SaveAs(newFilePath, XlFileFormat.xlWorkbookDefault);
                newWorkbook.Application.DisplayAlerts = true;
                newWorkbook.Close();
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("\n\n"+ex.ToString());
                newWorkbook.Close(false);
                excelApp.Quit();
                return -1;
            }

            return 0;
        }

        static bool RowsEqual(Range row1, Range row2)
        {
            for (int i = 1; i <= row1.Columns.Count; i++)
            {
                if (!Equals(row1.Cells[1, i].Value, row2.Cells[1, i].Value))
                {
                    return false;
                }
            }
            return true;
        }

        static bool HasExcessColumns(Range usedRange, int expectedNumberOfColumns)
        {
            for (int col = expectedNumberOfColumns + 1; col <= usedRange.Columns.Count; col++)
            {
                for (int row = 2; row <= usedRange.Rows.Count; row++)
                {
                    if (usedRange.Cells[row, col].Value != null)
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        private static Workbook CreateExcelWorkbookFromTextFile(Application excel, string filePath)
        {
            // Create a new workbook object
            Workbook workbook = excel.Workbooks.Add();

            // Add a new worksheet to the workbook object
            Worksheet worksheet = workbook.Worksheets.Add();

            // Create a connection string to the text file using Power Query
            string connectionString = $"TEXT;{filePath}";

            try
            {
                QueryTable queryTable = worksheet.QueryTables.Add(Connection: connectionString, Destination: worksheet.Cells[1, 1]);
                queryTable.TextFileParseType = XlTextParsingType.xlDelimited;
                queryTable.HasAutoFormat = true;
                queryTable.PreserveFormatting = true;
                queryTable.TextFileTextQualifier = XlTextQualifier.xlTextQualifierNone;
                queryTable.Refresh();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }

            return workbook;
        }

        private static Workbook CreateExcelWorkbookFromTextFile2(Application excel, string filePath)
        {
            // Create a new workbook object
            Workbook workbook = excel.Workbooks.Add();

            // Add a new worksheet to the workbook object
            Worksheet worksheet = workbook.Worksheets.Add();

            // Create a connection string to the text file using Power Query
            string connectionString = $"TEXT;{filePath}";

            try
            {
                QueryTable queryTable = worksheet.QueryTables.Add(Connection: connectionString, Destination: worksheet.Cells[1, 1]);
                //queryTable.TextFileParseType = XlTextParsingType.xlDelimited;
                /*queryTable.HasAutoFormat = true;
                queryTable.PreserveFormatting = true;*/
                queryTable.HasAutoFormat = false;
                queryTable.PreserveFormatting = false;
                //queryTable.TextFileTextQualifier = XlTextQualifier.xlTextQualifierNone;
                queryTable.Refresh();
                
                // Specify the column data types to treat columns containing numbers with zeros at the front as text
                int numColumns = queryTable.ResultRange.Columns.Count;
                for (int i = 1; i <= numColumns; i++)
                {
                    Range column = queryTable.ResultRange.Columns[i];
                    for (int j = 1; j <= column.Cells.Count; j++)
                        if (IsNumericWithLeadingZeros(column.Cells[j].Value))
                        {
                            column.NumberFormat = "@";
                            queryTable.Refresh();
                            break;
                        }
                }

                queryTable.Refresh();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }

            return workbook;
        }

        // Helper function to check if a value is a numeric string with leading zeros
        private static bool IsNumericWithLeadingZeros(object value)
        {
            if (value == null || !(value is string))
            {
                return false;
            }

            string strValue = (string)value;
            if (strValue.Length == 0)
            {
                return false;
            }

            if (!char.IsDigit(strValue[0]) && strValue[0] != '-')
            {
                return false;
            }

            if (strValue[0] == '0' && strValue.Length > 1)
            {
                return true;
            }

            double result;
            return double.TryParse(strValue, out result);
        }

        private static string[] LaunchFilePicker()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Excel Workbook (*.xlsx;*.xlsb;*.xlsm;*.xls;*.txt)|*.xlsx;*.xlsb;*.xlsm;*.xls;*.txt";
            openFileDialog.Title = "Select Excel workbooks or text files to merge";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                return openFileDialog.FileNames;
            }
            else
            {
                return null;
            }
        }

        public static string PromptForSaveLocation()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx|Excel Binary Workbook (*.xlsb)|*.xlsb|Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm|Excel 97-2003 Workbook (*.xls)|*.xls";
            saveFileDialog.Title = "Save Merged Workbook";
            saveFileDialog.ShowDialog();

            if (saveFileDialog.FileName != "")
            {
                return saveFileDialog.FileName;
            }
            else
            {
                throw new Exception("No file selected.");
            }
        }
    }
}