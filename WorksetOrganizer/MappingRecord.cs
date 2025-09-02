using System;
using System.Text.RegularExpressions;

namespace WorksetOrchestrator
{
    public class MappingRecord
    {
        public string WorksetName { get; set; }
        public string SystemNameInModel { get; set; }
        public string SystemDescription { get; set; }
        public string ModelPackageCode { get; set; }

        public string NormalizedPackageCode
        {
            get
            {
                if (ModelPackageCode == "NO EXPORT")
                    return "NO EXPORT";

                // Normalize by stripping all digits and 'x' placeholders, keep letters and hyphens
                return Regex.Replace(ModelPackageCode ?? "", @"[\dx\s]+", "");
            }
        }

        /// <summary>
        /// Gets a simplified pattern for better matching (removes xx, xxx placeholders)
        /// </summary>
        public string SimplifiedSystemPattern
        {
            get
            {
                if (string.IsNullOrEmpty(SystemNameInModel))
                    return string.Empty;

                // Remove common placeholder patterns
                string simplified = SystemNameInModel
                    .Replace("xxx", "")
                    .Replace("xx", "")
                    .Replace("x", "")
                    .Trim();

                return simplified;
            }
        }

        public override string ToString()
        {
            return $"System: '{SystemNameInModel}' → Workset: '{WorksetName}' → Package: '{ModelPackageCode}'";
        }
    }
}