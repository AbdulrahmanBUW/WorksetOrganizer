using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;

namespace WorksetOrchestrator
{
    public static class ExcelReader
    {
        public static List<MappingRecord> ReadMapping(string excelFilePath)
        {
            var mappingList = new List<MappingRecord>();

            if (!File.Exists(excelFilePath))
                throw new FileNotFoundException($"Excel file not found: {excelFilePath}");

            // EPPlus 4.x doesn't need LicenseContext - remove this line completely
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                    throw new Exception("Excel file contains no worksheets.");

                // Try to find the 'Mapping' worksheet, or use the first one
                ExcelWorksheet worksheet = null;

                // Case-insensitive search for 'Mapping' worksheet
                foreach (var ws in package.Workbook.Worksheets)
                {
                    if (ws.Name.Equals("Mapping", StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet = ws;
                        break;
                    }
                }

                // If no 'Mapping' worksheet found, use the first one
                if (worksheet == null)
                    worksheet = package.Workbook.Worksheets[1]; // EPPlus worksheets are 1-indexed

                if (worksheet == null)
                    throw new Exception("No accessible worksheet found in Excel file.");

                if (worksheet.Dimension == null)
                    throw new Exception("The worksheet appears to be empty.");

                int rowCount = worksheet.Dimension.End.Row;
                int colCount = worksheet.Dimension.End.Column;

                if (rowCount < 2)
                    throw new Exception("Excel file must contain at least a header row and one data row.");

                // Find header indices
                var headers = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int col = 1; col <= colCount; col++)
                {
                    var cellValue = worksheet.Cells[1, col].Value;
                    string headerValue = cellValue?.ToString()?.Trim();
                    if (!string.IsNullOrEmpty(headerValue))
                        headers[headerValue] = col;
                }

                // Required headers (case-insensitive)
                string[] requiredHeaders = { "Workset Name", "System Name in Model File", "System Description", "Model iFLS/Package Code" };
                var missingHeaders = new List<string>();

                foreach (string header in requiredHeaders)
                {
                    if (!headers.ContainsKey(header))
                        missingHeaders.Add(header);
                }

                if (missingHeaders.Count > 0)
                    throw new Exception($"Required headers not found: {string.Join(", ", missingHeaders)}. Available headers: {string.Join(", ", headers.Keys)}");

                // Read data rows
                for (int row = 2; row <= rowCount; row++)
                {
                    var record = new MappingRecord
                    {
                        WorksetName = GetCellValue(worksheet, row, headers["Workset Name"]),
                        SystemNameInModel = GetCellValue(worksheet, row, headers["System Name in Model File"]),
                        SystemDescription = GetCellValue(worksheet, row, headers["System Description"]),
                        ModelPackageCode = GetCellValue(worksheet, row, headers["Model iFLS/Package Code"])
                    };

                    // Only add records with valid workset name and system name (unless it's a special case)
                    if (!string.IsNullOrEmpty(record.WorksetName) &&
                        (!string.IsNullOrEmpty(record.SystemNameInModel) || record.ModelPackageCode == "NO EXPORT"))
                    {
                        mappingList.Add(record);
                    }
                }

                if (mappingList.Count == 0)
                    throw new Exception("No valid mapping records found in Excel file.");
            }

            return mappingList;
        }

        private static string GetCellValue(ExcelWorksheet worksheet, int row, int col)
        {
            try
            {
                var cellValue = worksheet.Cells[row, col].Value;
                return cellValue?.ToString()?.Trim();
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}