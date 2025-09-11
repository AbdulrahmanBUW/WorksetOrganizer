using System;
using System.IO;
using System.Reflection;
using Autodesk.Revit.UI;
using System.Windows.Media;
using System.Windows; // Add this for FlowDirection
using System.Globalization; // Add this for CultureInfo
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
            string assemblyDirectory = Path.GetDirectoryName(assemblyPath);

            // Create push button
            PushButtonData buttonData = new PushButtonData(
                "WorksetOrganiser",
                "Workset" + Environment.NewLine + "Organiser",
                assemblyPath,
                "WorksetOrchestrator.Command");

            buttonData.ToolTip = "Organizes worksets based on Excel mapping and export RVTs.";

            // Try to set icons with better error handling
            try
            {
                // Method 1: Try embedded resources first
                buttonData.LargeImage = GetEmbeddedImage("WorksetOrganizer.Icons.WorksetOrganizer_32x32.png");
                buttonData.Image = GetEmbeddedImage("WorksetOrganizer.Icons.WorksetOrganizer_16x16.png");

                // Method 2: If embedded resources fail, try file paths
                if (buttonData.LargeImage == null || buttonData.Image == null)
                {
                    string iconPath32 = Path.Combine(assemblyDirectory, "Icons", "WorksetOrganizer_32x32.png");
                    string iconPath16 = Path.Combine(assemblyDirectory, "Icons", "WorksetOrganizer_16x16.png");

                    if (File.Exists(iconPath32) && buttonData.LargeImage == null)
                    {
                        buttonData.LargeImage = new System.Windows.Media.Imaging.BitmapImage(new Uri(iconPath32, UriKind.Absolute));
                    }

                    if (File.Exists(iconPath16) && buttonData.Image == null)
                    {
                        buttonData.Image = new System.Windows.Media.Imaging.BitmapImage(new Uri(iconPath16, UriKind.Absolute));
                    }
                }

                // Method 3: Use a simple built-in icon as fallback
                if (buttonData.LargeImage == null && buttonData.Image == null)
                {
                    // Create a simple colored rectangle as fallback
                    buttonData.LargeImage = CreateFallbackIcon(32);
                    buttonData.Image = CreateFallbackIcon(16);
                }
            }
            catch (Exception ex)
            {
                // Log the error for debugging
                System.Diagnostics.Debug.WriteLine($"Icon loading failed: {ex.Message}");

                // Use fallback icon
                buttonData.LargeImage = CreateFallbackIcon(32);
                buttonData.Image = CreateFallbackIcon(16);
            }

            // Add button to panel
            PushButton button = panel.AddItem(buttonData) as PushButton;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        // Helper method to load embedded images
        private System.Windows.Media.ImageSource GetEmbeddedImage(string resourceName)
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                Stream stream = assembly.GetManifestResourceStream(resourceName);

                if (stream != null)
                {
                    var bitmap = new System.Windows.Media.Imaging.BitmapImage();
                    bitmap.BeginInit();
                    bitmap.StreamSource = stream;
                    bitmap.EndInit();
                    return bitmap;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to load embedded image {resourceName}: {ex.Message}");
            }

            return null;
        }

        // Create a simple fallback icon
        private System.Windows.Media.ImageSource CreateFallbackIcon(int size)
        {
            try
            {
                var drawingVisual = new System.Windows.Media.DrawingVisual();
                using (var context = drawingVisual.RenderOpen())
                {
                    // Draw a simple colored rectangle with "WO" text
                    var brush = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 120, 215)); // Blue color
                    var rect = new System.Windows.Rect(0, 0, size, size);
                    context.DrawRectangle(brush, null, rect);

                    // Add text
                    var formattedText = new System.Windows.Media.FormattedText(
                        "WO",
                        CultureInfo.GetCultureInfo("en-us"),
                        FlowDirection.LeftToRight,
                        new System.Windows.Media.Typeface("Arial"),
                        size * 0.4,
                        System.Windows.Media.Brushes.White,
                        1.0);

                    var textPoint = new System.Windows.Point(
                        (size - formattedText.Width) / 2,
                        (size - formattedText.Height) / 2);

                    context.DrawText(formattedText, textPoint);
                }

                var renderTargetBitmap = new System.Windows.Media.Imaging.RenderTargetBitmap(
                    size, size, 96, 96, System.Windows.Media.PixelFormats.Pbgra32);
                renderTargetBitmap.Render(drawingVisual);

                return renderTargetBitmap;
            }
            catch
            {
                return null;
            }
        }
    }
}