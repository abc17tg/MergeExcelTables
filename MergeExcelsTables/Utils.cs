using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace MergeExcelsTables
{
    public static class Utils
    {
        public static List<string> ExcelExt = new List<string> { ".xlsx", ".xlsb", "xlsm", "xls" };
        public static List<string> TextExt = new List<string> { ".txt", ".csv" };

        public static object[,] ConvertTo2DArray(List<object[]> list)
        {
            int rows = list.Count;
            int cols = list[0].Length;
            object[,] array = new object[rows, cols];

            for (int i = 0; i < rows; i++)
            {
                for (int j = 0; j < cols; j++)
                {
                    array[i, j] = list[i][j];
                }
            }

            return array;
        }

        public static T[,] ConvertTo2DArray2<T>(IList<T[]> arrays)
        {
            // TODO: Validation and special-casing for arrays.Count == 0
            int minorLength = arrays[0].Length;
            T[,] ret = new T[arrays.Count, minorLength];
            for (int i = 0; i < arrays.Count; i++)
            {
                var array = arrays[i];
                if (array.Length != minorLength)
                {
                    throw new ArgumentException
                        ("All arrays must be the same length");
                }
                for (int j = 0; j < minorLength; j++)
                {
                    ret[i, j] = array[j];
                }
            }
            return ret;
        }

        public static string LaunchFilePickerSingle()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Excel Workbook (*.xlsx;*.xlsb;*.xlsm;*.xls;*.txt;*.csv)|*.xlsx;*.xlsb;*.xlsm;*.xls;*.txt;*.csv";
            openFileDialog.Title = "Select Excel workbook or text file as template and to merge";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return null;

            return openFileDialog.FileName;
        }

        public static List<string> LaunchFilePicker()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Excel Workbook (*.xlsx;*.xlsb;*.xlsm;*.xls;*.txt;*.csv)|*.xlsx;*.xlsb;*.xlsm;*.xls;*.txt;*.csv";
            openFileDialog.Title = "Select rest of Excel workbooks and text files to merge (order will be random)";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return null;

            var filePaths = openFileDialog.FileNames.ToList();
            var tempFilePaths = filePaths.Where(p => ExcelExt.Contains(Path.GetExtension(p).ToLower())).ToList();
            tempFilePaths.AddRange(filePaths.Where(p => TextExt.Contains(Path.GetExtension(p).ToLower())).ToList());
            filePaths = tempFilePaths;
            return filePaths;
        }

        public static string PromptForSaveLocation()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx|Excel Binary Workbook (*.xlsb)|*.xlsb|Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm|Excel 97-2003 Workbook (*.xls)|*.xls";
            saveFileDialog.Title = "Save Merged Workbook";
            saveFileDialog.DefaultExt = "xlsx";
            saveFileDialog.AddExtension = true;
            string downloadsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Downloads");
            saveFileDialog.InitialDirectory = string.IsNullOrEmpty(downloadsPath) ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) : downloadsPath;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.ShowDialog();

            if (string.IsNullOrEmpty(saveFileDialog.FileName))
                throw new Exception("No file selected.");

            return saveFileDialog.FileName;
        }

        public static bool SaveAs(Excel.Workbook workbook)
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
