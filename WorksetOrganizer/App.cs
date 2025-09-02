using System;
using System.IO;
using System.Reflection;
using Autodesk.Revit.UI;
using OfficeOpenXml; // EPPlus

namespace WorksetOrchestrator
{
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            // Create a custom ribbon tab
            string tabName = "Workset Tools";
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch { } // Tab might already exist

            // Create a ribbon panel
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Organization");

            // Get assembly path
            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            // Create push button
            PushButtonData buttonData = new PushButtonData(
                "WorksetOrganiser",
                "Workset" + Environment.NewLine + "Organiser",
                assemblyPath,
                "WorksetOrchestrator.Command");

            buttonData.ToolTip = "Organizes worksets based on Excel mapping and export RVTs.";

            // Add button to panel
            PushButton button = panel.AddItem(buttonData) as PushButton;

            return Result.Succeeded;
        }


        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
