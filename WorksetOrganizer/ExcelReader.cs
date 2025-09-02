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

            // Set EPPlus license context (required for EPPlus 5.0+)
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                if (package.Workbook.Worksheets.Count == 0)
                    throw new Exception("Excel file contains no worksheets.");

                // Try to find the 'Mapping' worksheet, or use the first one
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Mapping"] ??
                                         package.Workbook.Worksheets[0];

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
                    string headerValue = worksheet.Cells[1, col].Text?.Trim();
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
                        WorksetName = worksheet.Cells[row, headers["Workset Name"]].Text?.Trim(),
                        SystemNameInModel = worksheet.Cells[row, headers["System Name in Model File"]].Text?.Trim(),
                        SystemDescription = worksheet.Cells[row, headers["System Description"]].Text?.Trim(),
                        ModelPackageCode = worksheet.Cells[row, headers["Model iFLS/Package Code"]].Text?.Trim()
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
    }
}