using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;

namespace WorksetOrchestrator
{
    public static class ExcelReader
    {
        public static List<MappingRecord> ReadMapping(string excelFilePath)
        {
            var mappingList = new List<MappingRecord>();
            Application excelApp = null;
            Workbook workbook = null;

            try
            {
                excelApp = new Application { Visible = false };
                workbook = excelApp.Workbooks.Open(excelFilePath);

                Worksheet worksheet = workbook.Sheets["Mapping"] as Worksheet;
                if (worksheet == null)
                    throw new Exception("Worksheet named 'Mapping' not found.");

                Range usedRange = worksheet.UsedRange;
                int rowCount = usedRange.Rows.Count;
                int colCount = usedRange.Columns.Count;

                // Find column indices by header name
                var headers = new Dictionary<string, int>();
                for (int col = 1; col <= colCount; col++)
                {
                    string headerValue = (usedRange.Cells[1, col] as Range)?.Value2?.ToString();
                    if (!string.IsNullOrEmpty(headerValue))
                        headers[headerValue.Trim()] = col;
                }

                // Verify all required headers are present
                string[] requiredHeaders = { "Workset Name", "System Name in Model File", "System Description", "Model iFLS/Package Code" };
                foreach (string header in requiredHeaders)
                {
                    if (!headers.ContainsKey(header))
                        throw new Exception($"Required header '{header}' not found in Excel sheet.");
                }

                // Read data rows
                for (int row = 2; row <= rowCount; row++)
                {
                    var record = new MappingRecord
                    {
                        WorksetName = GetCellValue(usedRange, row, headers["Workset Name"]),
                        SystemNameInModel = GetCellValue(usedRange, row, headers["System Name in Model File"]),
                        SystemDescription = GetCellValue(usedRange, row, headers["System Description"]),
                        ModelPackageCode = GetCellValue(usedRange, row, headers["Model iFLS/Package Code"])
                    };

                    // Only add if it has essential data
                    if (!string.IsNullOrEmpty(record.WorksetName) && !string.IsNullOrEmpty(record.SystemNameInModel))
                        mappingList.Add(record);
                }
            }
            finally
            {
                // Clean up COM objects
                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
            }

            return mappingList;
        }

        private static string GetCellValue(Range range, int row, int col)
        {
            try
            {
                return (range.Cells[row, col] as Range)?.Value2?.ToString() ?? "";
            }
            catch
            {
                return "";
            }
        }
    }
}