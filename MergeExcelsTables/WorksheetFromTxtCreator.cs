using System;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Data;

namespace MergeExcelsTables
{
    public static class WorksheetFromTxtCreator
    {
        public static Excel.Workbook CreateExcelWorkbookFromTextFileQueryTable(Application excel, string filePath, string delimiter = "\t")
        {
            // Create a new workbook object
            Excel.Workbook workbook = excel.Workbooks.Add();

            // Add a new worksheet to the workbook object
            Excel.Worksheet worksheet = workbook.Worksheets.Add();

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

            return workbook;
        }

        public static Excel.Workbook CreateExcelWorkbookFromTextFile(Excel.Application excel, string filePath, string delimiter)
        {
            // Create a new workbook object
            Excel.Workbook workbook = excel.Workbooks.Add();

            // Add a new worksheet to the workbook object
            Excel.Worksheet worksheet = workbook.Worksheets.Add();

            // Load the text file into a DataTable
            DataTable dataTable = LoadTextFileToDataTable(filePath, delimiter);

            // Write the contents of the DataTable to the Excel worksheet
            DataTableToExcelWorksheet(dataTable, worksheet);

            return workbook;
        }

        public static DataTable LoadTextFileToDataTable(string filePath, string delimiter)
        {
            DataTable dataTable = new DataTable();

            using (StreamReader sr = new StreamReader(filePath))
            {
                string[] headers = sr.ReadLine().Split(new string[] { delimiter }, StringSplitOptions.None);

                foreach (string header in headers)
                {
                    dataTable.Columns.Add(header);
                }

                while (!sr.EndOfStream)
                {
                    string[] rows = sr.ReadLine().Split(new string[] { delimiter }, StringSplitOptions.None);
                    DataRow newRow = dataTable.NewRow();

                    for (int i = 0; i < rows.Length; i++)
                    {
                        newRow[i] = rows[i];
                    }

                    dataTable.Rows.Add(newRow);
                }
            }

            return dataTable;
        }

        
        public static void DataTableToExcelWorksheet(DataTable dataTable, Excel.Worksheet worksheet)
        {
            // Write column headers
            var headers = dataTable.Columns.Cast<DataColumn>()
                .Select((c, i) => new { Column = c, Index = i + 1 })
                .ToList();

            var headerRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, headers.Count]];
            headerRange.Value2 = headers.Select(h => h.Column.ColumnName).ToArray();

            // Write data rows
            var data = dataTable.Rows.Cast<DataRow>()
                .SelectMany(r => headers.Select(h => r[h.Index - 1]))
                .ToArray();

            var dataRange = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[data.Length / headers.Count + 1, headers.Count]];
            dataRange.Value = data;

            // Set number format
            dataRange.NumberFormat = "@";
        }
    }
}


/*private static void DataTableToExcelWorksheet(DataTable dataTable, Excel.Worksheet worksheet)
        {
            // Write column headers
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName;
            }

            // Write data rows
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                for (int j = 0; j < dataTable.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1].NumberFormat = "@";
                    worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j];
                }
            }
        }*/

/* public static void DataTableToExcelWorksheet(DataTable dataTable, Excel.Worksheet worksheet)
         {
             // Write column headers
             var headers = dataTable.Columns.Cast<DataColumn>()
                 .Select((c, i) => new { Column = c, Index = i + 1 })
                 .ToList();

             var headerRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, headers.Count]];
             headerRange.Value2 = headers.Select(h => h.Column.ColumnName).ToArray();

             // Write data rows
             var data = dataTable.Rows.Cast<DataRow>()
                 .Select(r => headers.Select(h => r[h.Index - 1]).ToArray())
                 .ToList();

             var dataRange = worksheet.Range[worksheet.Cells[2, 1], worksheet.Cells[data.Count + 1, headers.Count]];
             dataRange.Value2 = Utils.ConvertTo2DArray2(data);

             dataRange.NumberFormat = "@";
         }*/

