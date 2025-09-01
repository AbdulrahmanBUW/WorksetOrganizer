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

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var worksheet = package.Workbook.Worksheets["Mapping"];
                if (worksheet == null)
                    throw new Exception("Worksheet named 'Mapping' not found.");

                int rowCount = worksheet.Dimension.End.Row;
                int colCount = worksheet.Dimension.End.Column;

                // Find header indices
                var headers = new Dictionary<string, int>();
                for (int col = 1; col <= colCount; col++)
                {
                    string headerValue = worksheet.Cells[1, col].Text?.Trim();
                    if (!string.IsNullOrEmpty(headerValue))
                        headers[headerValue] = col;
                }

                string[] requiredHeaders = { "Workset Name", "System Name in Model File", "System Description", "Model iFLS/Package Code" };
                foreach (string header in requiredHeaders)
                    if (!headers.ContainsKey(header))
                        throw new Exception($"Required header '{header}' not found in Excel sheet.");

                // Read data
                for (int row = 2; row <= rowCount; row++)
                {
                    var record = new MappingRecord
                    {
                        WorksetName = worksheet.Cells[row, headers["Workset Name"]].Text,
                        SystemNameInModel = worksheet.Cells[row, headers["System Name in Model File"]].Text,
                        SystemDescription = worksheet.Cells[row, headers["System Description"]].Text,
                        ModelPackageCode = worksheet.Cells[row, headers["Model iFLS/Package Code"]].Text
                    };

                    if (!string.IsNullOrEmpty(record.WorksetName) && !string.IsNullOrEmpty(record.SystemNameInModel))
                        mappingList.Add(record);
                }
            }

            return mappingList;
        }
    }
}
