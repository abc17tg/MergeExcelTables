using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;

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

        public static string[] LaunchFilePicker()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "Excel Workbook (*.xlsx;*.xlsb;*.xlsm;*.xls;*.txt;*.csv)|*.xlsx;*.xlsb;*.xlsm;*.xls;*.txt;*.csv";
            openFileDialog.Title = "Select Excel workbooks or text files to merge";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return null;

            var filePaths = openFileDialog.FileNames;
            var tempFilePaths = filePaths.Where(p => ExcelExt.Contains(Path.GetExtension(p).ToLower())).ToList();
            tempFilePaths.AddRange(filePaths.Where(p => TextExt.Contains(Path.GetExtension(p).ToLower())).ToList());
            filePaths = tempFilePaths.ToArray();
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

    }
}
