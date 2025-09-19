using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace CustomRevitCommand
{
    /// <summary>
    /// Command to open dimension settings dialog
    /// </summary>
    [Transaction(TransactionMode.Manual)]
    public class DimensionSettingsCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Load current settings
                var currentSettings = DimensionSettings.LoadFromProject(doc);

                // Show settings dialog
                var settingsWindow = new SettingsWindow(doc, currentSettings);
                bool? result = settingsWindow.ShowDialog();

                if (result == true)
                {
                    // Settings were saved in the dialog
                    TaskDialog.Show("Settings", "Dimension settings have been saved successfully.\n\n" +
                        "The new settings will be used for all future auto-dimension operations in this project.");
                    return Result.Succeeded;
                }
                else
                {
                    // User cancelled
                    return Result.Cancelled;
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Error opening settings dialog: {ex.Message}");
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}