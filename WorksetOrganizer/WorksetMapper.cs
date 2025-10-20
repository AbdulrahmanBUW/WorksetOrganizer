using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace WorksetOrchestrator
{
    public class WorksetMapper
    {
        private readonly Dictionary<string, string> _worksetToIflsMapping;
        private readonly Dictionary<string, string> _packageToIflsMapping;

        public WorksetMapper()
        {
            _worksetToIflsMapping = InitializeWorksetToIflsMapping();
            _packageToIflsMapping = InitializePackageToIflsMapping();
        }

        private Dictionary<string, string> InitializeWorksetToIflsMapping()
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                {"DX_BDA", "B-D"},
                {"DX_CDA", "D-D"},
                {"DX_CHM", "C-S"},
                {"DX_CKE", "C-L"},
                {"DX_ELT", "E-S"},
                {"DX_EXH", "A-X"},
                {"DX_PAW", "S-D"},
                {"DX_PG", "G-B"},
                {"DX_PKW", "P-D"},
                {"DX_PS", "G-S"},
                {"DX_PWI", "U-D"},
                {"DX_VAC", "V-D"},
                {"DX_SLUR", "M-S"},
                {"DX_STB", "STB"},
                {"DX_UPW", "U-D"},
                {"DX_PVAC", "V-V"},
                {"DX_RR", "RR"},
                {"DX_FND", "FND"},
                {"DX_Sub-tool", "S-BT"},
                {"DX_Tool", "T-L"}
            };
        }

        private Dictionary<string, string> InitializePackageToIflsMapping()
        {
            return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                {"B-D", "B-D"},
                {"D-D", "D-D"},
                {"C-S", "C-S"},
                {"C-L", "C-L"},
                {"E-S", "E-S"},
                {"ELT", "E-S"},
                {"A-X", "A-X"},
                {"S-D", "S-D"},
                {"G-B", "G-B"},
                {"P-D", "P-D"},
                {"G-S", "G-S"},
                {"U-D", "U-D"},
                {"V-D", "V-D"},
                {"M-S", "M-S"},
                {"V-V", "V-V"},
                {"STB", "STB"}
            };
        }

        public string GetIflsCodeForWorkset(string worksetName)
        {
            if (_worksetToIflsMapping.TryGetValue(worksetName, out string iflsCode))
            {
                return iflsCode;
            }

            return worksetName.Replace("DX_", "").Substring(0, Math.Min(3, worksetName.Replace("DX_", "").Length));
        }

        public string GetIflsCodeForPackage(string packageKey)
        {
            if (packageKey.Equals("QC", StringComparison.OrdinalIgnoreCase))
                return "QC";

            if (_packageToIflsMapping.TryGetValue(packageKey, out string iflsCode))
            {
                return iflsCode;
            }

            string cleanedKey = packageKey.Replace("4xx", "").Replace("xxx", "").Trim();
            return cleanedKey;
        }

        public bool IsWorksetCategoryBased(string worksetName)
        {
            if (string.IsNullOrEmpty(worksetName))
                return false;

            var categoryBasedWorksets = new[]
            {
                "DX_STB",
                "DX_RR",
                "DX_FND",
                "DX_ELT"
            };

            return Array.Exists(categoryBasedWorksets, w =>
                w.Equals(worksetName, StringComparison.OrdinalIgnoreCase));
        }
    }
}


