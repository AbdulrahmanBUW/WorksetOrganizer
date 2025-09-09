using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

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

        // Template integration parameters
        private bool _isTemplateIntegration;
        private List<string> _extractedFiles;
        private string _templateFilePath;

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
            _isTemplateIntegration = false;
            _extractedFiles = null;
            _templateFilePath = null;
        }

        public void SetTemplateIntegrationParameters(WorksetOrchestrator orchestrator,
            List<string> extractedFiles, string templateFilePath, string destinationPath)
        {
            _orchestrator = orchestrator;
            _extractedFiles = extractedFiles;
            _templateFilePath = templateFilePath;
            _destinationPath = destinationPath;
            _isComplete = false;
            _success = false;
            _lastException = null;
            _isTemplateIntegration = true;
            _mapping = null;
            _overwriteFiles = false;
            _exportQc = false;
        }

        public void Execute(UIApplication app)
        {
            try
            {
                if (_isTemplateIntegration)
                {
                    // When doing template integration, orchestrator will handle SafeOpenDocument internally
                    _orchestrator.LogMessage($"Event handler: starting template integration for {_extractedFiles?.Count ?? 0} files.");
                    _success = _orchestrator.IntegrateIntoTemplate(_extractedFiles, _templateFilePath, _destinationPath);
                }
                else
                {
                    _orchestrator.LogMessage("Event handler: starting standard Execute.");
                    _success = _orchestrator.Execute(_mapping, _destinationPath, _overwriteFiles, _exportQc);
                }
            }
            catch (Exception ex)
            {
                _success = false;
                _lastException = ex;
                // Log the error through the orchestrator
                _orchestrator?.LogMessage($"ERROR in event handler: {ex.Message}");
                _orchestrator?.LogMessage($"Stack Trace: {ex.StackTrace}");
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

    /// <summary>
    /// Handler for duplicate type names during copy operations
    /// Automatically handles duplicate types without showing dialogs
    /// </summary>
    public class DuplicateTypeNamesHandler : IDuplicateTypeNamesHandler
    {
        public DuplicateTypeAction OnDuplicateTypeNamesFound(DuplicateTypeNamesHandlerArgs args)
        {
            // Automatically use existing types from destination document
            return DuplicateTypeAction.UseDestinationTypes;
        }
    }
}
