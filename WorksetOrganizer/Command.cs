using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace WorksetOrchestrator
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;

                if (uiDoc == null)
                {
                    TaskDialog.Show("Error", "Please open a document first.");
                    return Result.Failed;
                }

                IntPtr revitHandle = uiApp.MainWindowHandle;

                MainForm mainForm = new MainForm(uiDoc, revitHandle);
                mainForm.Show();

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                message = ex.ToString();
                return Result.Failed;
            }
        }
    }
}