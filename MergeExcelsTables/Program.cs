using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Data;
using WTC = MergeExcelsTables.WorksheetFromTxtCreator;
using System.Collections.Generic;

namespace MergeExcelsTables
{
    internal class Program
    {
        private static bool m_showPrompts = true;
        private static Application m_excelApp;
        private static Excel.Workbook m_workbook;
        private static Excel.Workbook m_newWorkbook;

        [STAThread]
        static int Main(string[] args)
        {
            List<string> filePaths;
            if (args.Length > 0 && (args.Last() == "false" || args.Last() == "true"))
            {
                if (args.Last() == "false" || args.Last() == "true")
                {
                    m_showPrompts = bool.TryParse(args.Last(), out bool result);
                    if (!result)
                        m_showPrompts = true;
                }
            }

            filePaths = args.Where(p => File.Exists(p)).ToList();

            if (filePaths.Count < 2)
                filePaths = GetPaths();

            m_excelApp = new Application { Visible = false, CutCopyMode = 0, DisplayAlerts = false };
            m_workbook = m_excelApp.Workbooks.Add();
            m_newWorkbook = m_excelApp.Workbooks.Add();

            try
            {
                bool firstFile = true;
                int currentRow = 2;
                string delimiter = "\t";

                if (m_showPrompts && filePaths.Any(p=>Utils.TextExt.Contains(Path.GetExtension(p))))
                {
                    // Prompt the user for the delimiter using the DelimiterForm
                    var delimiterForm = new DelimiterForm();
                    DialogResult result = delimiterForm.ShowDialog();
                    delimiterForm.Activate();
                    if (result == DialogResult.OK)
                        delimiter = delimiterForm.Delimiter;
                }

                foreach (string filePath in filePaths)
                {
                    Console.WriteLine($"Processing: ./{new DirectoryInfo(Path.GetDirectoryName(filePath)).Name}/{Path.GetFileName(filePath)}");

                    Merge(filePath, firstFile, delimiter, ref currentRow);                 
                    firstFile = false;

                    Console.SetCursorPosition(0, Console.CursorTop - 1);
                    Console.Write(new string(' ', Console.WindowWidth));
                    Console.SetCursorPosition(0, Console.CursorTop-1);
                    Console.WriteLine($"Processed: ./{new DirectoryInfo(Path.GetDirectoryName(filePath)).Name}/{Path.GetFileName(filePath)}");
                }

                FormatRange(m_newWorkbook.Worksheets[1].UsedRange);

                Utils.SaveAs(m_newWorkbook);

                m_excelApp.Workbooks.Close();
                m_excelApp.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine("\n\n" + ex.ToString());
                m_excelApp.Workbooks.Close();
                m_excelApp.Quit();
                return -1;
            }

            return 0;
        }

        private static void FormatRange(Excel.Range range)
        {
            // Format the range as a table with headers
            Excel.ListObject table = m_newWorkbook.Worksheets[1].ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange,
                range, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing);
            table.Name = "Table1";
            table.TableStyle = "TableStyleLight1";
        }

        private static void Merge(string filePath, bool firstFile, string delimiter, ref int currentRow)
        {

            string fileExt = Path.GetExtension(filePath).ToLower();
            Excel.Range usedRange;

            if (Utils.TextExt.Contains(fileExt))
            {
                //usedRange = WTC.CreateExcelWorkbookFromTextFile(m_workbook.Worksheets.Add(), filePath, delimiter);
                usedRange = WTC.CreateExcelWorkbookFromTextFileQueryTable(m_workbook.Worksheets.Add(), filePath, delimiter);
            }
            else if (Utils.ExcelExt.Contains(fileExt))
                usedRange = m_excelApp.Workbooks.Open(filePath).Worksheets[1].usedRange;
            else
                throw new Exception();

            if (m_workbook == null)
                throw new Exception();

            // Check if the first row is the same and if there are any excess columns
            Range headerRow = usedRange.Rows[1];
            int lastColumnNumber = (int)m_excelApp.WorksheetFunction.CountA(headerRow.Rows[1]);
            headerRow = usedRange.Range["A1", usedRange.Cells[1, lastColumnNumber]];

            if (firstFile)
                headerRow.Copy(m_newWorkbook.Worksheets[1].Rows[1]);

            else if (!RowsEqual(headerRow, m_newWorkbook.Worksheets[1].Rows[1]))
                throw new Exception("Headers don't match each other\n");

            if (HasExcessColumns(usedRange, lastColumnNumber))
                throw new Exception("Headers don't match columns count\n");

            // Copy the used range below the first row to the new workbook
            Excel.Range newRange = usedRange.Offset[1, 0].Resize[usedRange.Rows.Count - 1];
            if (!firstFile && currentRow > 2)
            {
                Range firstRow = m_newWorkbook.Worksheets[1].UsedRange.Rows[3];
                for (int i = 1; i <= firstRow.Cells.Count; i++)
                {
                    m_newWorkbook.Worksheets[1].Columns[i].NumberFormat = firstRow.Cells[i].NumberFormat;
                }
            }

            newRange.Copy();
            if (firstFile)
                m_newWorkbook.Worksheets[1].Cells[currentRow, 1].PasteSpecial(Excel.XlPasteType.xlPasteFormulasAndNumberFormats);
            else
                m_newWorkbook.Worksheets[1].Cells[currentRow, 1].PasteSpecial(Excel.XlPasteType.xlPasteFormulas);

            currentRow += usedRange.Rows.Count - 1;
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

        private static List<string> GetPaths()
        {
            // Launch the file picker dialog to allow the user to select files
            string firstFilePath = Utils.LaunchFilePickerSingle();
            if (string.IsNullOrEmpty(firstFilePath))
            {
                Console.WriteLine("Error: No template file selected.");
                return null;
            }
            List<string> filePaths = Utils.LaunchFilePicker();
            if (filePaths == null || filePaths.Count < 1)
            {
                Console.WriteLine("Error: No files selected.");
                return null;
            }
            filePaths.Remove(firstFilePath);
            filePaths.Insert(0, firstFilePath);
            return filePaths;
        }
    }
}