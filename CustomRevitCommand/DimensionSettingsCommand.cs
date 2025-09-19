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
                    // Get the updated settings
                    var newSettings = settingsWindow.Settings;

                    // Create detailed feedback message
                    string settingsInfo = "Dimension settings have been saved successfully!\n\n";
                    settingsInfo += "Current Settings:\n";
                    settingsInfo += $"• Reference Type: {GetReferenceTypeDisplay(newSettings.ReferenceType)}\n";
                    settingsInfo += $"• Include Grids: {(newSettings.IncludeGrids ? "Yes" : "No")}\n";
                    settingsInfo += $"• Include Levels: {(newSettings.IncludeLevels ? "Yes" : "No")}\n";
                    settingsInfo += $"• Include Structural: {(newSettings.IncludeStructural ? "Yes" : "No")}\n";
                    settingsInfo += $"• Include Curtain Walls: {(newSettings.IncludeCurtainWalls ? "Yes" : "No")}\n";
                    settingsInfo += $"• Include Curtain Wall Grids: {(newSettings.IncludeMullions ? "Yes" : "No")}\n";
                    settingsInfo += $"• Default Offset: {newSettings.DefaultOffset:F2} feet\n\n";

                    // Add reference type explanation
                    switch (newSettings.ReferenceType)
                    {
                        case DimensionReferenceType.Centerline:
                            settingsInfo += "Dimensions will be created to element centerlines (traditional method).";
                            break;
                        case DimensionReferenceType.ExteriorFace:
                            settingsInfo += "Dimensions will be created to exterior faces of walls (best for building exterior).";
                            break;
                        case DimensionReferenceType.InteriorFace:
                            settingsInfo += "Dimensions will be created to interior faces of walls (best for room layouts).";
                            break;
                        case DimensionReferenceType.Auto:
                            settingsInfo += "System will automatically choose the best reference type for each element.";
                            break;
                    }

                    settingsInfo += "\n\nThese settings will be used for all Auto-Dimension and Dimension Chain operations.";

                    TaskDialog settingsDialog = new TaskDialog("Settings Saved");
                    settingsDialog.MainContent = settingsInfo;
                    settingsDialog.CommonButtons = TaskDialogCommonButtons.Ok;
                    settingsDialog.MainIcon = TaskDialogIcon.TaskDialogIconInformation;
                    settingsDialog.Show();

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

        /// <summary>
        /// Get display-friendly name for reference type
        /// </summary>
        private string GetReferenceTypeDisplay(DimensionReferenceType referenceType)
        {
            switch (referenceType)
            {
                case DimensionReferenceType.Centerline:
                    return "Centerline";
                case DimensionReferenceType.ExteriorFace:
                    return "Exterior Face";
                case DimensionReferenceType.InteriorFace:
                    return "Interior Face";
                case DimensionReferenceType.Auto:
                    return "Auto (Smart Selection)";
                default:
                    return "Unknown";
            }
        }
    }
}