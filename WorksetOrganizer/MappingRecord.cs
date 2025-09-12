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

                return Regex.Replace(ModelPackageCode ?? "", @"[\dx\s]+", "");
            }
        }

        public string SimplifiedSystemPattern
        {
            get
            {
                if (string.IsNullOrEmpty(SystemNameInModel))
                    return string.Empty;

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