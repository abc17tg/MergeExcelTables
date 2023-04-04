using System;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Collections.Generic;

namespace MergeExcelsTables
{
    public static class WorksheetFromTxtCreator
    {
        public static Excel.Range CreateExcelWorkbookFromTextFileQueryTable(Excel.Worksheet worksheet, string filePath, string delimiter = "\t")
        {
            // Create a connection string to the text file using Power Query
            string connectionString = $"TEXT;{filePath}";

            try
            {
                Excel.QueryTable queryTable = worksheet.QueryTables.Add(Connection: connectionString, Destination: worksheet.Cells[1, 1]);
                queryTable.TextFileParseType = Excel.XlTextParsingType.xlDelimited;
                queryTable.TextFileOtherDelimiter = delimiter;
                queryTable.HasAutoFormat = true;
                queryTable.PreserveFormatting = true;
                queryTable.TextFileTextQualifier = Excel.XlTextQualifier.xlTextQualifierNone;
                queryTable.Refresh();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message, ex);
            }
            return worksheet.UsedRange;
        }

        /*public static Excel.Range CreateExcelWorkbookFromTextFile(Excel.Worksheet worksheet, string filePath, string delimiter = "\t")
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"File not found at path: {filePath}");
            }

            using (StreamReader reader = new StreamReader(filePath))
            {
                List<List<string>> rows = new List<List<string>>();

                string line;

                while ((line = reader.ReadLine()) != null)
                {
                    string[] values = line.Split(new string[] { delimiter }, StringSplitOptions.None);

                    rows.Add(values.ToList());
                }

                int rowCount = rows.Count;
                int columnCount = rows.Max(row => row.Count);

                Excel.Range range = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[rowCount, columnCount]];
                range.Value2 = rows.Select(row => row.ToArray()).ToArray();
                range.NumberFormat = "@";

                return range;
            }
        }*/
    }
}

