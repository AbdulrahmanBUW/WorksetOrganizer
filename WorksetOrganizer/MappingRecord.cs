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

                // Normalize by stripping all digits and 'x' placeholders
                return Regex.Replace(ModelPackageCode ?? "", @"[\dx]", "");
            }
        }
    }
}
