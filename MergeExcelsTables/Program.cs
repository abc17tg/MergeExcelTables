using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Data;
using WTC = MergeExcelsTables.WorksheetFromTxtCreator;
using System.Threading.Tasks;

namespace MergeExcelsTables
{
    internal class Program
    {
        private static bool m_showPrompts = true;

        [STAThread]
        static int Main(string[] args)
        {
            string[] filePaths;
            if (args.Length > 0 && (args.Last() == "false" || args.Last() == "true"))
            {
                if (args.Last() == "false" || args.Last() == "true")
                {
                    m_showPrompts = bool.TryParse(args.Last(), out bool result);
                    if (!result)
                        m_showPrompts = true;
                }
            }

            filePaths = args.Where(p => File.Exists(p)).ToArray();

            if (filePaths.Length < 2)
            {
                // Launch the file picker dialog to allow the user to select files
                filePaths = Utils.LaunchFilePicker();
                if (filePaths == null)
                {
                    Console.WriteLine("Error: No files selected.");
                    return -1;
                }
            }

            Application excelApp = new Application { Visible = false, CutCopyMode = 0, DisplayAlerts = false };
            Excel.Workbook newWorkbook = excelApp.Workbooks.Add();

            try
            {
                bool firstFile = true;
                int currentRow = 2;
                foreach (string filePath in filePaths)
                {
                    Console.WriteLine($"Processing: {filePath}");
                    string fileExt = Path.GetExtension(filePath).ToLower();
                    Excel.Workbook workbook;

                    if (Utils.TextExt.Contains(fileExt))
                    {
                        string delimiter = "\t";

                        if (firstFile && m_showPrompts)
                        {
                            // Prompt the user for the delimiter using the DelimiterForm
                            var delimiterForm = new DelimiterForm();
                            DialogResult result = delimiterForm.ShowDialog();
                            delimiter = delimiterForm.Delimiter;
                            workbook = WTC.CreateExcelWorkbookFromTextFileQueryTable(excelApp, filePath, delimiter);
                        }
                        else// if (firstFile)
                            workbook = WTC.CreateExcelWorkbookFromTextFileQueryTable(excelApp, filePath);
                        /*else
                            workbook = WTC.CreateExcelWorkbookFromTextFile(excelApp, filePath, delimiter);*/
                    }
                    else if (Utils.ExcelExt.Contains(fileExt))
                        workbook = excelApp.Workbooks.Open(filePath);
                    else
                        throw new Exception();

                    if (workbook == null)
                        throw new Exception();

                    // Check if the first row is the same and if there are any excess columns
                    Range usedRange = workbook.Worksheets[1].UsedRange;
                    Range headerRow = usedRange.Rows[1];
                    int lastColumnNumber = (int)excelApp.WorksheetFunction.CountA(headerRow.Rows[1]);
                    headerRow = usedRange.Range["A1", usedRange.Cells[1, lastColumnNumber]];

                    if (firstFile)
                        headerRow.Copy(newWorkbook.Worksheets[1].Rows[1]);

                    else if (!RowsEqual(headerRow, newWorkbook.Worksheets[1].Rows[1]))
                        throw new Exception();

                    if (HasExcessColumns(usedRange, lastColumnNumber))
                        throw new Exception();

                    // Copy the used range below the first row to the new workbook
                    Excel.Range newRange = usedRange.Offset[1, 0].Resize[usedRange.Rows.Count - 1];
                    if (!firstFile && currentRow > 2)
                    {
                        Range firstRow = newWorkbook.Worksheets[1].UsedRange.Rows[3];
                        for (int i = 1; i <= firstRow.Cells.Count; i++)
                        {
                            newWorkbook.Worksheets[1].Columns[i].NumberFormat = firstRow.Cells[i].NumberFormat;
                        }
                    }

                    newRange.Copy();
                    if (firstFile)
                    {
                        newWorkbook.Worksheets[1].Cells[currentRow, 1].PasteSpecial(Excel.XlPasteType.xlPasteFormulasAndNumberFormats);
                        firstFile = false;
                    }
                    else
                        newWorkbook.Worksheets[1].Cells[currentRow, 1].PasteSpecial(Excel.XlPasteType.xlPasteFormulas);

                    currentRow += usedRange.Rows.Count - 1;
                    
                    Console.SetCursorPosition(0, Console.CursorTop - 1);
                    Console.WriteLine($"Processed: {filePath}");

                    Task.Factory.StartNew(() => { workbook.Close(false); });
                }

                Range finalRange = newWorkbook.Worksheets[1].UsedRange;
                // Format the range as a table with headers
                Excel.ListObject table = newWorkbook.Worksheets[1].ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange,
                    finalRange, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing);
                table.Name = "MyTable";
                table.TableStyle = "TableStyleLight1";

                SaveAs(newWorkbook);
                excelApp.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("\n\n" + ex.ToString());
                excelApp.Quit();
                return -1;
            }

            return 0;
        }

        private static bool SaveAs(Excel.Workbook workbook)
        {
            try
            {
                string newFilePath = Utils.PromptForSaveLocation();
                if (!string.IsNullOrEmpty(newFilePath))
                {
                    Excel.XlFileFormat fileFormat;
                    string fileExtension = Path.GetExtension(newFilePath);

                    switch (fileExtension)
                    {
                        case ".xlsx":
                            fileFormat = Excel.XlFileFormat.xlOpenXMLWorkbook;
                            break;
                        case ".xlsm":
                            fileFormat = Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled;
                            break;
                        case ".xls":
                            fileFormat = Excel.XlFileFormat.xlExcel8;
                            break;
                        case ".xlsb":
                            fileFormat = Excel.XlFileFormat.xlExcel12;
                            break;
                        default:
                            fileFormat = Excel.XlFileFormat.xlWorkbookDefault;
                            break;
                    }

                    workbook.SaveAs(newFilePath, fileFormat, ReadOnlyRecommended: false);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine("\n\n" + ex.ToString());
                return false;
            }

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

        private static bool IsNumericWithLeadingZeros(object value)
        {
            if (value == null || !(value is string))
                return false;

            string strValue = (string)value;
            if (strValue.Length == 0)
                return false;

            if (!char.IsDigit(strValue[0]) && strValue[0] != '-')
                return false;

            if (strValue[0] == '0' && strValue.Length > 1)
                return true;

            double result;
            return double.TryParse(strValue, out result);
        }


    }
}