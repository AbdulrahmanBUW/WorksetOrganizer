using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;

namespace WorksetOrchestrator
{
    public class WorksetEventHandler : IExternalEventHandler
    {
        private WorksetOrchestrator _orchestrator;
        private List<MappingRecord> _mapping;
        private string _destinationPath;
        private bool _overwriteFiles;
        private bool _exportQc;
        private bool _isComplete;
        private bool _success;
        private Exception _lastException;

        public bool IsComplete => _isComplete;
        public bool Success => _success;
        public Exception LastException => _lastException;

        public void SetParameters(WorksetOrchestrator orchestrator, List<MappingRecord> mapping,
            string destinationPath, bool overwriteFiles, bool exportQc)
        {
            _orchestrator = orchestrator;
            _mapping = mapping;
            _destinationPath = destinationPath;
            _overwriteFiles = overwriteFiles;
            _exportQc = exportQc;
            _isComplete = false;
            _success = false;
            _lastException = null;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                _success = _orchestrator.Execute(_mapping, _destinationPath, _overwriteFiles, _exportQc);
            }
            catch (Exception ex)
            {
                _success = false;
                _lastException = ex;
                // Log the error through the orchestrator
                _orchestrator.LogMessage($"ERROR in event handler: {ex.Message}");
                _orchestrator.LogMessage($"Stack Trace: {ex.StackTrace}");
            }
            finally
            {
                _isComplete = true;
            }
        }

        public string GetName()
        {
            return "WorksetOrchestratorEventHandler";
        }
    }
}