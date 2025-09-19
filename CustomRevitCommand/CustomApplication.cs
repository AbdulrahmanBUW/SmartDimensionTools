using System;
using System.IO;
using System.Windows.Media.Imaging;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;

namespace CustomRevitCommand
{
    // External Application class - handles add-in initialization
    public class CustomApplication : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            string tabName = "DEAXO Draw";
            try
            {
                a.CreateRibbonTab(tabName);
            }
            catch
            {
                // Tab might already exist
            }

            RibbonPanel ribbonPanel = a.CreateRibbonPanel(tabName, "Smart Dimensions");
            string thisAssemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;

            // First button: Auto-Dimension (Enhanced)
            PushButtonData buttonData = new PushButtonData(
                "AutoDimensionCommand",
                "Auto-\nDimension",
                thisAssemblyPath,
                "CustomRevitCommand.AutoDimensionCommand"
            );
            buttonData.ToolTip = "Creates dimension chains including grids, levels, elements, and curtain walls\n\n" +
                                "Features:\n" +
                                "• Multiple view support\n" +
                                "• Curtain wall and mullion dimensioning\n" +
                                "• Face/centerline reference options\n" +
                                "• Adaptive tolerances\n\n" +
                                "Hold Ctrl when clicking to open settings";
            buttonData.LongDescription = "Advanced auto-dimensioning tool with support for:\n\n" +
                                        "• Standard building elements (walls, beams, columns)\n" +
                                        "• Curtain walls and mullion centerlines\n" +
                                        "• Grids and levels\n" +
                                        "• Multiple reference types (centerline, exterior face, interior face)\n" +
                                        "• Batch processing across multiple views\n\n" +
                                        "The tool intelligently groups collinear elements and creates " +
                                        "dimension chains with optimal placement.";

            // Add icon to the button
            try
            {
                string iconPath = Path.Combine(Path.GetDirectoryName(thisAssemblyPath), "Icons");
                if (Directory.Exists(iconPath))
                {
                    string largeIcon = Path.Combine(iconPath, "AutoDimension32.png");
                    string smallIcon = Path.Combine(iconPath, "AutoDimension16.png");

                    if (File.Exists(largeIcon))
                        buttonData.LargeImage = new BitmapImage(new Uri(largeIcon));

                    if (File.Exists(smallIcon))
                        buttonData.Image = new BitmapImage(new Uri(smallIcon));
                }
            }
            catch (Exception ex)
            {
                // If icon loading fails, continue without icon
                System.Diagnostics.Debug.WriteLine($"Could not load icon: {ex.Message}");
            }

            PushButton pushButton = ribbonPanel.AddItem(buttonData) as PushButton;

            // Add separator
            ribbonPanel.AddSeparator();

            // Second button: Dimension Chain (Enhanced)
            PushButtonData chainButtonData = new PushButtonData(
                "DimensionChainCommand",
                "Dimension\nChain",
                thisAssemblyPath,
                "CustomRevitCommand.DimensionChainCommand"
            );
            chainButtonData.ToolTip = "Create dimension chains by defining direction line and placement point\n\n" +
                                     "Interactive tool:\n" +
                                     "• Pick start and end points\n" +
                                     "• Automatic element detection\n" +
                                     "• Continuous mode until ESC\n" +
                                     "• Supports all element types";
            chainButtonData.LongDescription = "Interactive dimension chain creation tool:\n\n" +
                                             "1. Click to pick start point\n" +
                                             "2. Click to pick end point\n" +
                                             "3. Tool finds all intersecting elements\n" +
                                             "4. Creates dimension chain automatically\n" +
                                             "5. Continues until ESC is pressed\n\n" +
                                             "Works with all element types including curtain walls and mullions.";

            // Add icon for chain command
            try
            {
                string iconPath = Path.Combine(Path.GetDirectoryName(thisAssemblyPath), "Icons");
                if (Directory.Exists(iconPath))
                {
                    string largeIcon = Path.Combine(iconPath, "DimensionChain32.png");
                    string smallIcon = Path.Combine(iconPath, "DimensionChain16.png");

                    if (File.Exists(largeIcon))
                        chainButtonData.LargeImage = new BitmapImage(new Uri(largeIcon));

                    if (File.Exists(smallIcon))
                        chainButtonData.Image = new BitmapImage(new Uri(smallIcon));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Could not load chain icon: {ex.Message}");
            }

            PushButton chainButton = ribbonPanel.AddItem(chainButtonData) as PushButton;

            // Add separator
            ribbonPanel.AddSeparator();

            // Third button: Settings
            PushButtonData settingsButtonData = new PushButtonData(
                "DimensionSettingsCommand",
                "Dimension\nSettings",
                thisAssemblyPath,
                "CustomRevitCommand.DimensionSettingsCommand"
            );
            settingsButtonData.ToolTip = "Configure Auto-Dimension settings\n\n" +
                                        "Configure:\n" +
                                        "• Reference types (face/centerline)\n" +
                                        "• Element inclusion options\n" +
                                        "• Tolerances and spacing\n" +
                                        "• Behavior preferences";
            settingsButtonData.LongDescription = "Open the settings dialog to configure:\n\n" +
                                                "• Dimension reference type (centerline, exterior face, interior face)\n" +
                                                "• Which element types to include (grids, levels, structural, curtain walls)\n" +
                                                "• Collinearity and perpendicular tolerances\n" +
                                                "• Default offsets and spacing\n" +
                                                "• Progress dialog and behavior options\n\n" +
                                                "Settings are saved per project and user.";

            // Add icon for settings command
            try
            {
                string iconPath = Path.Combine(Path.GetDirectoryName(thisAssemblyPath), "Icons");
                if (Directory.Exists(iconPath))
                {
                    string largeIcon = Path.Combine(iconPath, "Settings32.png");
                    string smallIcon = Path.Combine(iconPath, "Settings16.png");

                    if (File.Exists(largeIcon))
                        settingsButtonData.LargeImage = new BitmapImage(new Uri(largeIcon));

                    if (File.Exists(smallIcon))
                        settingsButtonData.Image = new BitmapImage(new Uri(smallIcon));
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Could not load settings icon: {ex.Message}");
            }

            PushButton settingsButton = ribbonPanel.AddItem(settingsButtonData) as PushButton;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}