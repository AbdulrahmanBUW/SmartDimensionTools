using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

// Alias to resolve namespace conflicts
using RevitGrid = Autodesk.Revit.DB.Grid;



namespace CustomRevitCommand
{
    // External Command class - Auto Dimension (Multi-View) with all improvements
    [Transaction(TransactionMode.Manual)]
    public class AutoDimensionCommand : IExternalCommand
    {
        // Settings-driven constants
        private DimensionSettings _settings;
        private double PARALLEL_TOLERANCE => 0.05; // ~3 degrees for 2D
        private double PERPENDICULAR_TOLERANCE => _settings?.PerpendicularTolerance ?? 0.1;
        private double DIMENSION_OFFSET => _settings?.DefaultOffset ?? 1.64; // feet
        private string _logFilePath;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Load settings first
                _settings = DimensionSettings.LoadFromProject(doc);

                // Show settings dialog if user wants to configure
                var settingsResult = ShowSettingsDialogIfNeeded(doc);
                if (settingsResult == Result.Cancelled)
                {
                    return Result.Cancelled;
                }

                // STEP 1: Check if any elements are selected
                ICollection<ElementId> selectedIds = uidoc.Selection.GetElementIds();

                if (selectedIds.Count == 0)
                {
                    TaskDialog.Show("Error", "Please select at least one element before running the Auto-Dimension tool.");
                    return Result.Cancelled;
                }

                // STEP 2: Show the WPF view selection dialog - PASS CURRENT VIEW
                ViewSelectionWindow viewSelectionWindow = new ViewSelectionWindow(doc, uidoc.ActiveView, _settings);
                bool? dialogResult = viewSelectionWindow.ShowDialog();

                // STEP 3: Check if user clicked OK and selected views
                if (dialogResult != true || viewSelectionWindow.SelectedViews.Count == 0)
                {
                    return Result.Cancelled; // User clicked Cancel or didn't select any views
                }

                // STEP 4: Get the selected views from the dialog
                List<View> selectedViews = viewSelectionWindow.SelectedViews;

                // STEP 5: Process each selected view with progress dialog
                return ProcessSelectedViews(doc, uidoc, selectedViews, selectedIds);
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Error", $"Exception: {ex.Message}");
                message = ex.Message;
                return Result.Failed;
            }
        }

        private void LogToFile(string message)
        {
            try
            {
                if (string.IsNullOrEmpty(_logFilePath))
                {
                    string tempPath = Path.GetTempPath();
                    _logFilePath = Path.Combine(tempPath, $"RevitAutoDimension_Debug_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                    File.WriteAllText(_logFilePath, $"Auto-Dimension Debug Log - {DateTime.Now:yyyy-MM-dd HH:mm:ss}\n");
                    File.AppendAllText(_logFilePath, "=".PadRight(80, '=') + "\n\n");
                }

                File.AppendAllText(_logFilePath, $"[{DateTime.Now:HH:mm:ss.fff}] {message}\n");
            }
            catch (Exception ex)
            {
                // Fallback - show in task dialog if file writing fails
                TaskDialog.Show("Debug Log Error", $"Could not write to log file: {ex.Message}\nOriginal message: {message}");
            }
        }


        /// <summary>
        /// Show settings dialog if needed (e.g., if Ctrl key is held)
        /// </summary>
        private Result ShowSettingsDialogIfNeeded(Document doc)
        {
            try
            {
                // Check if user wants to open settings (e.g., Ctrl key held down)
                bool showSettings = System.Windows.Forms.Control.ModifierKeys.HasFlag(System.Windows.Forms.Keys.Control);

                if (showSettings)
                {
                    var settingsWindow = new SettingsWindow(doc, _settings);
                    bool? result = settingsWindow.ShowDialog();

                    if (result == true)
                    {
                        _settings = settingsWindow.Settings;
                        return Result.Succeeded;
                    }
                    else
                    {
                        return Result.Cancelled;
                    }
                }

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error showing settings dialog: {ex.Message}");
                return Result.Succeeded; // Continue without settings
            }
        }

        /// <summary>
        /// Process all selected views with proper progress feedback
        /// </summary>
        private Result ProcessSelectedViews(Document doc, UIDocument uidoc, List<View> selectedViews, ICollection<ElementId> selectedIds)
        {
            int totalChainsCreated = 0;
            List<string> processedViews = new List<string>();
            List<string> failedViews = new List<string>();
            View originalActiveView = uidoc.ActiveView; // Remember the original active view

            try
            {
                using (TransactionGroup transGroup = new TransactionGroup(doc, "Auto-Dimension Multiple Views"))
                {
                    transGroup.Start();

                    for (int i = 0; i < selectedViews.Count; i++)
                    {
                        View view = selectedViews[i];

                        try
                        {
                            // STEP 6: Validate and change view OUTSIDE of transaction
                            if (!CanActivateView(view))
                            {
                                failedViews.Add($"✗ {view.Name}: Cannot activate view");
                                continue;
                            }

                            uidoc.ActiveView = view;

                            // Verify the switch actually worked
                            if (uidoc.ActiveView.Id != view.Id)
                            {
                                failedViews.Add($"✗ {view.Name}: Failed to activate view");
                                continue;
                            }

                            // STEP 7: Now start transaction for this view
                            using (Transaction trans = new Transaction(doc, $"Auto-Dimension - {view.Name}"))
                            {
                                trans.Start();

                                // STEP 8: Apply dimensioning logic to this view
                                int chainsCreated = ProcessViewForDimensioning(doc, view, selectedIds);

                                if (chainsCreated > 0)
                                {
                                    totalChainsCreated += chainsCreated;
                                    processedViews.Add($"✓ {view.Name}: {chainsCreated} chain(s)");
                                    trans.Commit();
                                }
                                else
                                {
                                    failedViews.Add($"⚠ {view.Name}: No dimensions created");
                                    trans.RollBack();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            failedViews.Add($"✗ {view.Name}: Error - {ex.Message}");
                            System.Diagnostics.Debug.WriteLine($"Error processing view {view.Name}: {ex.Message}");
                        }
                    }

                    // STEP 9: Restore original active view OUTSIDE of individual transactions
                    try
                    {
                        uidoc.ActiveView = originalActiveView;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Could not restore original view: {ex.Message}");
                        // This is not critical, just log it
                    }

                    // Complete the transaction group
                    if (totalChainsCreated > 0)
                    {
                        transGroup.Assimilate(); // Make it single undo operation
                    }
                    else
                    {
                        transGroup.RollBack();
                    }
                }

                // STEP 10: Show results summary to user
                ShowResultsSummary(totalChainsCreated, processedViews, failedViews);

                return totalChainsCreated > 0 ? Result.Succeeded : Result.Failed;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ProcessSelectedViews: {ex.Message}");
                return Result.Failed;
            }
        }

        /// <summary>
        /// Check if a view can be activated
        /// </summary>
        private bool CanActivateView(View view)
        {
            try
            {
                if (view == null) return false;
                if (view.IsTemplate) return false;

                // Check if view type supports dimensioning
                switch (view.ViewType)
                {
                    case ViewType.FloorPlan:
                    case ViewType.CeilingPlan:
                    case ViewType.AreaPlan:
                    case ViewType.Section:
                    case ViewType.Elevation:
                    case ViewType.Detail:
                        return true;
                    default:
                        return false;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error checking view activation: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Process a single view for dimensioning (view is already active when this is called)
        /// </summary>
        private int ProcessViewForDimensioning(Document doc, View view, ICollection<ElementId> selectedIds)
        {
            try
            {
                LogToFile($"=== PROCESSING VIEW: {view.Name} (Type: {view.ViewType}) ===");
                LogToFile($"Selected elements count: {selectedIds.Count}");

                // Log selected element IDs
                foreach (var id in selectedIds)
                {
                    var element = doc.GetElement(id);
                    LogToFile($"  Selected: {element?.Category?.Name} - {element?.Name} (ID: {id.IntegerValue})");
                }

                // STEP 1: Collect all items (elements, grids, levels, curtain walls)
                List<ProjectedItem> projectedItems = CollectAllItems(doc, view, selectedIds);

                if (projectedItems.Count == 0)
                {
                    LogToFile($"❌ No valid items found in view: {view.Name}");
                    LogToFile($"DIAGNOSIS: This is likely the main issue - no items are being collected");

                    // Show user the log file location
                    TaskDialog.Show("Debug - No Items Found",
                        $"No dimension items were collected.\n\n" +
                        $"Debug log saved to your Desktop:\n{Path.GetFileName(_logFilePath)}\n\n" +
                        $"Check the log file for detailed analysis.");
                    return 0;
                }

                LogToFile($"✅ Found {projectedItems.Count} items in view {view.Name}");

                // Log each projected item
                for (int i = 0; i < projectedItems.Count; i++)
                {
                    var item = projectedItems[i];
                    LogToFile($"  Item {i + 1}: {item.ItemType} - {item.Element?.Name ?? "Unnamed"} " +
                        $"(ID: {item.Element?.Id?.IntegerValue}) - Selected: {item.IsSelected}, IsPoint: {item.IsPointElement}");

                    if (!item.IsPointElement && item.ProjectedDirection != null)
                    {
                        LogToFile($"    Direction: ({item.ProjectedDirection.X:F3}, {item.ProjectedDirection.Y:F3}, {item.ProjectedDirection.Z:F3})");
                    }

                    if (item.ProjectedPoint != null)
                    {
                        LogToFile($"    Point: ({item.ProjectedPoint.X:F3}, {item.ProjectedPoint.Y:F3}, {item.ProjectedPoint.Z:F3})");
                    }
                }

                // STEP 2: Group by parallel directions with enhanced logic
                LogToFile($"\n--- GROUPING BY PARALLEL DIRECTIONS ---");
                var parallelGroups = GroupByParallelDirections(projectedItems);

                LogToFile($"Created {parallelGroups.Count} parallel groups:");
                foreach (var group in parallelGroups)
                {
                    var direction = group.Key;
                    var items = group.Value;
                    LogToFile($"  Group direction ({direction.X:F3}, {direction.Y:F3}, {direction.Z:F3}): {items.Count} items");

                    foreach (var item in items)
                    {
                        LogToFile($"    - {item.ItemType}: {item.Element?.Name ?? "Unnamed"} (Selected: {item.IsSelected})");
                    }
                }

                // STEP 3: Create combined dimension groups with semantic grouping
                LogToFile($"\n--- CREATING DIMENSION GROUPS ---");
                var dimensionGroups = CreateCombinedDimensionGroups(parallelGroups);

                LogToFile($"Created {dimensionGroups.Count} dimension groups:");
                foreach (var group in dimensionGroups)
                {
                    var direction = group.Key;
                    var items = group.Value;
                    LogToFile($"  Dimension group direction ({direction.X:F3}, {direction.Y:F3}, {direction.Z:F3}): {items.Count} items");

                    foreach (var item in items)
                    {
                        LogToFile($"    - {item.ItemType}: {item.Element?.Name ?? "Unnamed"} " +
                            $"(Pos: {item.PositionAlongDirection:F3})");
                    }
                }

                if (dimensionGroups.Count == 0)
                {
                    LogToFile($"❌ No dimension groups created - this means no valid groupings were found");
                    LogToFile($"DIAGNOSIS: Items were collected but couldn't be grouped for dimensioning");

                    TaskDialog.Show("Debug - No Groups",
                        $"Items were found but no dimension groups could be created.\n\n" +
                        $"Debug log saved to your Desktop:\n{Path.GetFileName(_logFilePath)}\n\n" +
                        $"Check the log file to see why grouping failed.");
                    return 0;
                }

                // STEP 4: Create dimension chains with proper reference types
                LogToFile($"\n--- CREATING DIMENSION CHAINS ---");
                int chainsCreated = 0;
                List<Dimension> createdDimensions = new List<Dimension>();

                foreach (var group in dimensionGroups)
                {
                    LogToFile($"Processing group with {group.Value.Count} items...");

                    if (group.Value.Count >= 2)
                    {
                        LogToFile($"  ✅ Sufficient items ({group.Value.Count}) for dimension chain");

                        try
                        {
                            Dimension newDimension = CreateDimensionChain(doc, view, group.Key, group.Value);
                            if (newDimension != null)
                            {
                                chainsCreated++;
                                createdDimensions.Add(newDimension);
                                LogToFile($"  ✅ Dimension chain created successfully! (Chain #{chainsCreated})");
                            }
                            else
                            {
                                LogToFile($"  ❌ CreateDimensionChain returned null");
                            }
                        }
                        catch (Exception ex)
                        {
                            LogToFile($"  ❌ Error creating dimension chain: {ex.Message}");
                            LogToFile($"     Stack trace: {ex.StackTrace}");
                        }
                    }
                    else
                    {
                        LogToFile($"  ❌ Insufficient items ({group.Value.Count}) for dimension chain (need at least 2)");
                    }
                }

                // STEP 5: Move dimensions for better placement
                if (createdDimensions.Count > 0)
                {
                    LogToFile($"\n--- MOVING DIMENSIONS ---");
                    try
                    {
                        MoveDimensionsLeft(doc, createdDimensions, 0.0328084); // 10mm in feet
                        LogToFile($"✅ Moved {createdDimensions.Count} dimensions");
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"❌ Error moving dimensions: {ex.Message}");
                    }
                }

                LogToFile($"\n=== PROCESSING COMPLETE: {chainsCreated} chains created ===");

                // Show final result to user
                if (chainsCreated == 0)
                {
                    TaskDialog.Show("Debug - No Dimensions Created",
                        $"No dimension chains were created.\n\n" +
                        $"Items found: {projectedItems.Count}\n" +
                        $"Groups created: {dimensionGroups.Count}\n\n" +
                        $"Full debug log saved to your Desktop:\n{Path.GetFileName(_logFilePath)}");
                }
                else
                {
                    TaskDialog.Show("Debug - Success!",
                        $"Created {chainsCreated} dimension chains!\n\n" +
                        $"Debug log saved to your Desktop:\n{Path.GetFileName(_logFilePath)}");
                }

                return chainsCreated;
            }
            catch (Exception ex)
            {
                LogToFile($"❌ Error processing view {view.Name}: {ex.Message}");
                LogToFile($"Stack trace: {ex.StackTrace}");

                TaskDialog.Show("Debug - Fatal Error",
                    $"Fatal error in processing:\n{ex.Message}\n\n" +
                    $"Debug log saved to your Desktop:\n{Path.GetFileName(_logFilePath)}");
                return 0;
            }
        }

        /// <summary>
        /// Show a summary of results to the user
        /// </summary>
        // REPLACE your ShowResultsSummary method in AutoDimensionCommand.cs with this version:
        private void ShowResultsSummary(int totalChains, List<string> processedViews, List<string> failedViews)
        {
            string summary = $"Auto-Dimension Results:\n\n";
            summary += $"Total dimension chains created: {totalChains}\n\n";

            if (processedViews.Count > 0)
            {
                summary += "Successfully processed views:\n";
                foreach (string view in processedViews)
                {
                    summary += $"{view}\n";
                }
                summary += "\n";
            }

            if (failedViews.Count > 0)
            {
                summary += "Views with issues:\n";
                foreach (string view in failedViews)
                {
                    summary += $"{view}\n";
                }
            }

            // Add settings information
            summary += $"\nSettings used:\n";
            summary += $"• Reference Type: {_settings.ReferenceType}\n";
            summary += $"• Include Curtain Walls: {(_settings.IncludeCurtainWalls ? "Yes" : "No")}\n";
            summary += $"• Include Mullions: {(_settings.IncludeMullions ? "Yes" : "No")}\n";

            // ADD CURTAIN WALL GRID LINE LIMITATION NOTICE:
            if (_settings.IncludeMullions)
            {
                summary += $"\nNote: Curtain Wall Grid Lines are not directly dimensionable in Revit.\n";
                summary += $"The tool dimensions to curtain walls and mullions instead.\n";
                summary += $"For precise curtain wall layout dimensions, consider:\n";
                summary += $"• Using the manual Dimension Chain tool\n";
                summary += $"• Dimensioning to actual mullion elements\n";
                summary += $"• Creating reference lines for custom dimensioning\n";
            }

            TaskDialog resultDialog = new TaskDialog("Auto-Dimension Results");
            resultDialog.MainContent = summary;
            resultDialog.CommonButtons = TaskDialogCommonButtons.Ok;

            if (totalChains > 0)
            {
                resultDialog.MainIcon = TaskDialogIcon.TaskDialogIconInformation;
            }
            else
            {
                resultDialog.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
            }

            resultDialog.Show();
        }

        // STEP 1: Collect all items (elements, grids, levels, perpendicular elements, curtain walls)
        private List<ProjectedItem> CollectAllItems(Document doc, View view, ICollection<ElementId> selectedIds)
        {
            List<ProjectedItem> projectedItems = new List<ProjectedItem>();

            try
            {
                LogToFile($"\n--- COLLECTING ITEMS FROM VIEW: {view.Name} ---");
                LogToFile($"View type: {view.ViewType}");
                LogToFile($"Selected IDs count: {selectedIds.Count}");

                // Get visible elements in view
                FilteredElementCollector collector = new FilteredElementCollector(doc, view.Id);
                var allViewElements = collector.ToElements().ToList();
                LogToFile($"Total elements in view: {allViewElements.Count}");

                int processedCount = 0;
                int skippedCount = 0;

                foreach (Element element in allViewElements)
                {
                    ProjectedItem projectedItem = null;
                    string skipReason = "";

                    try
                    {
                        if (element is RevitGrid grid && _settings.IncludeGrids)
                        {
                            LogToFile($"Processing grid: {grid.Name}");
                            projectedItem = ProcessGrid(grid, view);
                            if (projectedItem != null)
                            {
                                LogToFile($"  ✅ Grid processed successfully");
                            }
                            else
                            {
                                LogToFile($"  ❌ Grid processing failed");
                            }
                        }
                        else if (element is Level level && (view.ViewType == ViewType.Section || view.ViewType == ViewType.Elevation) && _settings.IncludeLevels)
                        {
                            LogToFile($"Processing level: {level.Name}");
                            projectedItem = ProcessLevel(level, view);
                            if (projectedItem != null)
                            {
                                LogToFile($"  ✅ Level processed successfully");
                            }
                            else
                            {
                                LogToFile($"  ❌ Level processing failed");
                            }
                        }
                        else if (HasCenterline(element))
                        {
                            bool isSelected = selectedIds.Contains(element.Id);
                            LogToFile($"Processing element: {element.Category?.Name} - {element.Name} (ID: {element.Id.IntegerValue}, Selected: {isSelected})");

                            if (isSelected)
                            {
                                // First try to process as regular linear element
                                projectedItem = ProcessElement(element, view);

                                if (projectedItem != null)
                                {
                                    LogToFile($"  ✅ Element processed as linear element");
                                }
                                else
                                {
                                    // If it fails as linear element, try as perpendicular element
                                    LogToFile($"  ⚠️ Linear processing failed, trying perpendicular...");
                                    projectedItem = ProcessPerpendicularElement(element, view);
                                    if (projectedItem != null)
                                    {
                                        LogToFile($"  ✅ Element processed as perpendicular element");
                                    }
                                    else
                                    {
                                        LogToFile($"  ❌ Both linear and perpendicular processing failed");
                                    }
                                }
                            }
                            else
                            {
                                skipReason = "Not selected";
                            }
                        }
                        else
                        {
                            skipReason = "No centerline or not included in settings";
                        }

                        if (projectedItem != null)
                        {
                            projectedItems.Add(projectedItem);
                            processedCount++;
                        }
                        else if (!string.IsNullOrEmpty(skipReason))
                        {
                            skippedCount++;
                            // Only log first few skipped items to avoid spam
                            if (skippedCount <= 5)
                            {
                                LogToFile($"Skipped: {element.Category?.Name} - {skipReason}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogToFile($"❌ Error processing element {element.Id}: {ex.Message}");
                    }
                }

                LogToFile($"View elements: Processed {processedCount}, Skipped {skippedCount}");

                // ENHANCED: Process curtain walls and their components with comprehensive grid support
                if (_settings.IncludeCurtainWalls || _settings.IncludeMullions)
                {
                    LogToFile($"\n--- PROCESSING CURTAIN WALL ELEMENTS ---");
                    LogToFile($"Include Curtain Walls: {_settings.IncludeCurtainWalls}");
                    LogToFile($"Include Mullions: {_settings.IncludeMullions}");

                    var curtainWallItems = CurtainWallHelper.ProcessCurtainWallElements(doc, view, selectedIds, _settings);
                    projectedItems.AddRange(curtainWallItems);
                    LogToFile($"✅ Added {curtainWallItems.Count} curtain wall elements");

                    // Log curtain wall items
                    foreach (var cwItem in curtainWallItems)
                    {
                        LogToFile($"  CW Item: {cwItem.ItemType} - {cwItem.Element?.Name ?? "Unnamed"} " +
                            $"(ID: {cwItem.Element?.Id?.IntegerValue}, Selected: {cwItem.IsSelected})");
                    }
                }

                // For layouts, also get ALL grids from project (not just visible ones)
                if (view.ViewType == ViewType.FloorPlan || view.ViewType == ViewType.CeilingPlan || view.ViewType == ViewType.AreaPlan)
                {
                    if (_settings.IncludeGrids)
                    {
                        LogToFile($"\n--- PROCESSING ALL PROJECT GRIDS ---");
                        FilteredElementCollector allGrids = new FilteredElementCollector(doc).OfClass(typeof(RevitGrid));
                        var allGridsList = allGrids.ToElements().Cast<RevitGrid>().ToList();
                        LogToFile($"Found {allGridsList.Count} total grids in project");

                        int addedGrids = 0;
                        foreach (RevitGrid grid in allGridsList)
                        {
                            // Skip if already processed
                            if (projectedItems.Any(p => p.Element.Id == grid.Id))
                            {
                                LogToFile($"Skipping already processed grid: {grid.Name}");
                                continue;
                            }

                            ProjectedItem projectedGrid = ProcessGrid(grid, view);
                            if (projectedGrid != null)
                            {
                                projectedItems.Add(projectedGrid);
                                addedGrids++;
                                LogToFile($"Added project grid: {grid.Name}");
                            }
                        }
                        LogToFile($"Added {addedGrids} additional grids from project");
                    }
                }

                LogToFile($"\n--- COLLECTION SUMMARY ---");
                LogToFile($"Total items collected: {projectedItems.Count}");

                // Summary by type
                var itemsByType = projectedItems.GroupBy(p => p.ItemType).ToDictionary(g => g.Key, g => g.Count());
                foreach (var kvp in itemsByType)
                {
                    LogToFile($"  {kvp.Key}: {kvp.Value}");
                }

                var curtainWallCount = projectedItems.Count(p => p.IsCurtainWallElement);
                var gridLineCount = projectedItems.Count(p => p.ItemType == "CurtainGridLine");
                var mullionCount = projectedItems.Count(p => p.ItemType == "Mullion");
                var selectedCount = projectedItems.Count(p => p.IsSelected);

                LogToFile($"  - Curtain wall elements: {curtainWallCount}");
                LogToFile($"  - Curtain grid lines: {gridLineCount}");
                LogToFile($"  - Mullions: {mullionCount}");
                LogToFile($"  - Selected items: {selectedCount}");

                LogToFile($"--- END COLLECTION ---\n");
            }
            catch (Exception ex)
            {
                LogToFile($"❌ Error collecting items: {ex.Message}");
                LogToFile($"Stack trace: {ex.StackTrace}");
            }

            return projectedItems;
        }

        private bool IsValidProjectedItem(ProjectedItem item)
        {
            if (item == null) return false;
            if (item.Element == null) return false;
            if (item.ProjectedPoint == null) return false;
            if (!item.IsPointElement && (item.ProjectedDirection == null || item.ProjectedDirection.GetLength() < 0.001)) return false;
            return true;
        }

        private string GetInvalidReason(ProjectedItem item)
        {
            if (item == null) return "Item is null";
            if (item.Element == null) return "Element is null";
            if (item.ProjectedPoint == null) return "ProjectedPoint is null";
            if (!item.IsPointElement && (item.ProjectedDirection == null || item.ProjectedDirection.GetLength() < 0.001))
                return "ProjectedDirection is null or too short for linear element";
            return "Unknown";
        }

        // Process elements that are perpendicular to the view surface - ENHANCED
        private ProjectedItem ProcessPerpendicularElement(Element element, View view)
        {
            try
            {
                if (!(element.Location is LocationCurve locationCurve)) return null;

                Curve curve = locationCurve.Curve;
                XYZ start3D = curve.GetEndPoint(0);
                XYZ end3D = curve.GetEndPoint(1);
                XYZ elementDirection3D = (end3D - start3D).Normalize();

                // Get view plane normal
                XYZ viewNormal = GetViewPlaneNormal(view);

                // Enhanced perpendicular detection
                double dotProduct = Math.Abs(elementDirection3D.DotProduct(viewNormal));
                bool isStrictlyPerpendicular = Math.Abs(dotProduct - 1.0) < PERPENDICULAR_TOLERANCE;

                // Also check if element is "mostly" perpendicular (within 30 degrees)
                bool isMostlyPerpendicular = dotProduct > 0.7; // cos(45°) ≈ 0.707

                if (!isStrictlyPerpendicular && !isMostlyPerpendicular) return null;

                System.Diagnostics.Debug.WriteLine($"Found perpendicular element: {element.Id} - {element.Category?.Name}");

                // Calculate intersection or center point
                XYZ projectedPoint;
                if (!isStrictlyPerpendicular)
                {
                    // For mostly perpendicular elements, calculate intersection with view plane
                    Line elementLine = Line.CreateBound(start3D, end3D);
                    XYZ intersectionPoint = CalculateViewPlaneIntersection(elementLine, view);
                    projectedPoint = intersectionPoint != null ?
                        ProjectPointToViewSurface(intersectionPoint, view) :
                        ProjectPointToViewSurface((start3D + end3D) / 2, view);
                }
                else
                {
                    // For strictly perpendicular, use center point
                    XYZ center3D = (start3D + end3D) / 2;
                    projectedPoint = ProjectPointToViewSurface(center3D, view);
                }

                Reference centerlineRef = GetCenterlineReference(element);

                return new ProjectedItem
                {
                    Element = element,
                    GeometricReference = centerlineRef,
                    CenterlineReference = centerlineRef,
                    ProjectedDirection = null, // No direction for point elements
                    ProjectedPoint = projectedPoint,
                    ItemType = "PerpendicularElement",
                    IsSelected = true,
                    IsPointElement = true,
                    ReferenceType = _settings.ReferenceType
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error processing perpendicular element: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Calculate intersection point of line with view plane
        /// </summary>
        private XYZ CalculateViewPlaneIntersection(Line elementLine, View view)
        {
            try
            {
                XYZ viewNormal = GetViewPlaneNormal(view);
                XYZ viewOrigin = view.Origin;

                // Line parameters
                XYZ lineStart = elementLine.GetEndPoint(0);
                XYZ lineDirection = elementLine.Direction;

                // Plane equation: (P - viewOrigin) • viewNormal = 0
                // Line equation: P = lineStart + t * lineDirection
                // Solve for t
                double denominator = lineDirection.DotProduct(viewNormal);

                if (Math.Abs(denominator) < 1e-10)
                {
                    // Line is parallel to plane
                    return null;
                }

                double t = (viewOrigin - lineStart).DotProduct(viewNormal) / denominator;

                // Check if intersection is within line bounds
                if (t >= 0 && t <= elementLine.Length)
                {
                    return lineStart + t * lineDirection;
                }

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error calculating view plane intersection: {ex.Message}");
                return null;
            }
        }

        // Get view plane normal vector
        private XYZ GetViewPlaneNormal(View view)
        {
            switch (view.ViewType)
            {
                case ViewType.FloorPlan:
                case ViewType.CeilingPlan:
                case ViewType.AreaPlan:
                    return XYZ.BasisZ; // Vertical elements are perpendicular to plan views

                case ViewType.Section:
                case ViewType.Elevation:
                    // For sections/elevations, the view direction is the normal to the cutting plane
                    return view.ViewDirection.Normalize();

                default:
                    return XYZ.BasisZ;
            }
        }

        private ProjectedItem ProcessGrid(RevitGrid grid, View view)
        {
            try
            {
                // For sections/elevations: Get view-specific geometry
                if (view.ViewType == ViewType.Section || view.ViewType == ViewType.Elevation)
                {
                    return ProcessGridInSectionView(grid, view);
                }
                else
                {
                    // For layouts: Use standard 3D projection
                    return ProcessGridInLayoutView(grid, view);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error processing grid {grid.Name}: {ex.Message}");
                return null;
            }
        }

        private ProjectedItem ProcessGridInSectionView(RevitGrid grid, View view)
        {
            try
            {
                // Get geometry as it actually appears in this view
                Options geometryOptions = new Options();
                geometryOptions.View = view;
                geometryOptions.IncludeNonVisibleObjects = false;

                GeometryElement geometryElement = grid.get_Geometry(geometryOptions);
                if (geometryElement == null) return null;

                Line gridLineInView = null;

                // Extract the line geometry from the grid as it appears in the view
                foreach (GeometryObject geoObj in geometryElement)
                {
                    if (geoObj is Line line)
                    {
                        gridLineInView = line;
                        break;
                    }
                    else if (geoObj is GeometryInstance geoInstance)
                    {
                        GeometryElement instGeometry = geoInstance.GetInstanceGeometry();
                        foreach (GeometryObject instObj in instGeometry)
                        {
                            if (instObj is Line instLine)
                            {
                                gridLineInView = instLine;
                                break;
                            }
                        }
                        if (gridLineInView != null) break;
                    }
                }

                if (gridLineInView == null) return null;

                // Use the line as it appears in the view
                XYZ start3D = gridLineInView.GetEndPoint(0);
                XYZ end3D = gridLineInView.GetEndPoint(1);

                // Convert to our 2D working space
                XYZ start2D = ProjectPointToViewSurface(start3D, view);
                XYZ end2D = ProjectPointToViewSurface(end3D, view);
                XYZ direction2D = (end2D - start2D);

                if (direction2D.GetLength() < 0.001) return null;

                direction2D = direction2D.Normalize();
                XYZ center2D = (start2D + end2D) / 2;

                return new ProjectedItem
                {
                    Element = grid,
                    GeometricReference = new Reference(grid),
                    CenterlineReference = new Reference(grid),
                    ProjectedDirection = direction2D,
                    ProjectedPoint = center2D,
                    ItemType = "Grid",
                    IsSelected = false,
                    IsPointElement = false,
                    ReferenceType = DimensionReferenceType.Centerline // Grids always use centerline
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error processing grid {grid.Name} in section: {ex.Message}");
                return null;
            }
        }

        private ProjectedItem ProcessGridInLayoutView(RevitGrid grid, View view)
        {
            try
            {
                if (!(grid.Curve is Line gridLine)) return null;

                // Get grid line points
                XYZ start3D = gridLine.GetEndPoint(0);
                XYZ end3D = gridLine.GetEndPoint(1);

                // Project to view surface
                XYZ start2D = ProjectPointToViewSurface(start3D, view);
                XYZ end2D = ProjectPointToViewSurface(end3D, view);

                XYZ direction2D = (end2D - start2D);
                if (direction2D.GetLength() < 0.001) return null;

                direction2D = direction2D.Normalize();
                XYZ center2D = (start2D + end2D) / 2;

                return new ProjectedItem
                {
                    Element = grid,
                    GeometricReference = new Reference(grid),
                    CenterlineReference = new Reference(grid),
                    ProjectedDirection = direction2D,
                    ProjectedPoint = center2D,
                    ItemType = "Grid",
                    IsSelected = false,
                    IsPointElement = false,
                    ReferenceType = DimensionReferenceType.Centerline
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error processing grid {grid.Name} in layout: {ex.Message}");
                return null;
            }
        }

        private ProjectedItem ProcessLevel(Level level, View view)
        {
            try
            {
                // For section/elevation views only
                if (view.ViewType != ViewType.Section && view.ViewType != ViewType.Elevation)
                    return null;

                // Try to get view-specific geometry first
                Options geometryOptions = new Options();
                geometryOptions.View = view;
                geometryOptions.IncludeNonVisibleObjects = false;

                GeometryElement geometryElement = level.get_Geometry(geometryOptions);

                if (geometryElement != null)
                {
                    // Look for level line in view-specific geometry
                    foreach (GeometryObject geoObj in geometryElement)
                    {
                        if (geoObj is Line levelLine)
                        {
                            XYZ start3D = levelLine.GetEndPoint(0);
                            XYZ end3D = levelLine.GetEndPoint(1);

                            // Convert to 2D working space
                            XYZ start2D = ProjectPointToViewSurface(start3D, view);
                            XYZ end2D = ProjectPointToViewSurface(end3D, view);
                            XYZ levelDirection2D = (end2D - start2D);

                            if (levelDirection2D.GetLength() > 0.001)
                            {
                                levelDirection2D = levelDirection2D.Normalize();
                                XYZ center2D = (start2D + end2D) / 2;

                                return new ProjectedItem
                                {
                                    Element = level,
                                    GeometricReference = new Reference(level),
                                    CenterlineReference = new Reference(level),
                                    ProjectedDirection = levelDirection2D,
                                    ProjectedPoint = center2D,
                                    ItemType = "Level",
                                    IsSelected = false,
                                    IsPointElement = false,
                                    ReferenceType = DimensionReferenceType.Centerline
                                };
                            }
                        }
                    }
                }

                // Fallback: Create level representation based on elevation
                double levelElevation = level.Elevation;
                XYZ levelPoint3D = new XYZ(0, 0, levelElevation);
                XYZ levelPoint2D = ProjectPointToViewSurface(levelPoint3D, view);
                XYZ fallbackDirection2D = new XYZ(1, 0, 0); // Horizontal in view

                return new ProjectedItem
                {
                    Element = level,
                    GeometricReference = new Reference(level),
                    CenterlineReference = new Reference(level),
                    ProjectedDirection = fallbackDirection2D,
                    ProjectedPoint = levelPoint2D,
                    ItemType = "Level",
                    IsSelected = false,
                    IsPointElement = false,
                    ReferenceType = DimensionReferenceType.Centerline
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error processing level {level.Name}: {ex.Message}");
                return null;
            }
        }

        private ProjectedItem ProcessElement(Element element, View view)
        {
            try
            {
                if (!(element.Location is LocationCurve locationCurve)) return null;

                Curve curve = locationCurve.Curve;
                XYZ start3D = curve.GetEndPoint(0);
                XYZ end3D = curve.GetEndPoint(1);

                // Project element centerline to view surface
                XYZ start2D = ProjectPointToViewSurface(start3D, view);
                XYZ end2D = ProjectPointToViewSurface(end3D, view);

                XYZ direction2D = (end2D - start2D);
                if (direction2D.GetLength() < 0.001) return null; // This will be caught as perpendicular element

                direction2D = direction2D.Normalize();
                XYZ center2D = (start2D + end2D) / 2;

                Reference centerlineRef = GetCenterlineReference(element);

                var projectedItem = new ProjectedItem
                {
                    Element = element,
                    GeometricReference = centerlineRef,
                    CenterlineReference = centerlineRef,
                    ProjectedDirection = direction2D,
                    ProjectedPoint = center2D,
                    ItemType = "Element",
                    IsSelected = true,
                    IsPointElement = false,
                    ReferenceType = _settings.ReferenceType
                };

                return projectedItem;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error processing element: {ex.Message}");
                return null;
            }
        }

        // Project 3D point to correct view surface
        private XYZ ProjectPointToViewSurface(XYZ point3D, View view)
        {
            switch (view.ViewType)
            {
                case ViewType.FloorPlan:
                case ViewType.CeilingPlan:
                case ViewType.AreaPlan:
                    return new XYZ(point3D.X, point3D.Y, 0);

                case ViewType.Section:
                    XYZ sectionRight = view.RightDirection.Normalize();
                    XYZ sectionUp = view.UpDirection.Normalize();
                    XYZ viewOrigin = view.Origin;
                    XYZ relativePoint = point3D - viewOrigin;

                    double x = relativePoint.DotProduct(sectionRight);
                    double y = relativePoint.DotProduct(sectionUp);
                    return new XYZ(x, y, 0);

                case ViewType.Elevation:
                    XYZ elevRight = view.RightDirection.Normalize();
                    XYZ elevUp = view.UpDirection.Normalize();
                    XYZ elevOrigin = view.Origin;
                    XYZ elevRelative = point3D - elevOrigin;

                    double elevX = elevRelative.DotProduct(elevRight);
                    double elevY = elevRelative.DotProduct(elevUp);
                    return new XYZ(elevX, elevY, 0);

                default:
                    return new XYZ(point3D.X, point3D.Y, 0);
            }
        }

        private bool HasCenterline(Element element)
        {
            if (element == null || element.Category == null) return false;
            if (element is RevitGrid || element is Level) return false;

            // Handle curtain wall elements
            if (CurtainWallHelper.IsCurtainWall(element as Wall))
            {
                return _settings.IncludeCurtainWalls;
            }

            // Handle curtain wall grid lines
            if (element is CurtainGridLine)
            {
                return _settings.IncludeMullions; // Curtain grid lines are controlled by mullion setting
            }

            // Handle mullions
            if (element is Mullion)
            {
                return _settings.IncludeMullions;
            }

            // Exclude fittings, accessories based on settings
            if (element is FamilyInstance)
            {
                string categoryName = element.Category.Name?.ToLower() ?? "";
                if (categoryName.Contains("fitting") ||
                    categoryName.Contains("accessory") ||
                    categoryName.Contains("insulation"))
                {
                    return _settings.IncludeStructural; // Only include if structural elements are enabled
                }
            }

            return element.Location is LocationCurve;
        }

        private Reference GetCenterlineReference(Element element)
        {
            try
            {
                if (element is FamilyInstance familyInstance)
                {
                    var centerlineRefs = familyInstance.GetReferences(FamilyInstanceReferenceType.CenterLeftRight);
                    if (centerlineRefs != null && centerlineRefs.Count > 0)
                        return centerlineRefs.FirstOrDefault();

                    var elevationRefs = familyInstance.GetReferences(FamilyInstanceReferenceType.CenterElevation);
                    if (elevationRefs != null && elevationRefs.Count > 0)
                        return elevationRefs.FirstOrDefault();
                }

                return new Reference(element);
            }
            catch
            {
                return new Reference(element);
            }
        }

        // STEP 2: Group by parallel directions - UPDATED to handle point elements with adaptive tolerances
        private Dictionary<XYZ, List<ProjectedItem>> GroupByParallelDirections(List<ProjectedItem> projectedItems)
        {
            var groups = new Dictionary<XYZ, List<ProjectedItem>>();

            System.Diagnostics.Debug.WriteLine($"=== GROUPING {projectedItems.Count} ITEMS BY PARALLEL DIRECTIONS ===");

            // Separate point elements and linear elements
            var pointElements = projectedItems.Where(item => item.IsPointElement).ToList();
            var linearElements = projectedItems.Where(item => !item.IsPointElement).ToList();

            System.Diagnostics.Debug.WriteLine($"Point elements: {pointElements.Count}, Linear elements: {linearElements.Count}");

            // Process linear elements first to establish direction groups
            foreach (var item in linearElements)
            {
                XYZ direction = item.ProjectedDirection;
                XYZ groupDirection = null;

                // Find existing parallel group
                foreach (var existingDirection in groups.Keys)
                {
                    if (AreDirectionsParallel(direction, existingDirection))
                    {
                        groupDirection = existingDirection;
                        break;
                    }
                }

                // Create new group if needed
                if (groupDirection == null)
                {
                    groupDirection = direction;
                    groups[groupDirection] = new List<ProjectedItem>();
                }

                groups[groupDirection].Add(item);
            }

            // Add point elements to ALL direction groups
            foreach (var pointElement in pointElements)
            {
                foreach (var group in groups.Values)
                {
                    group.Add(pointElement);
                }

                // If no groups exist yet, create a default horizontal group
                if (groups.Count == 0)
                {
                    XYZ defaultDirection = new XYZ(1, 0, 0);
                    groups[defaultDirection] = new List<ProjectedItem> { pointElement };
                }
            }

            return groups;
        }

        private bool AreDirectionsParallel(XYZ dir1, XYZ dir2)
        {
            double cross = Math.Abs(dir1.X * dir2.Y - dir1.Y * dir2.X);
            return cross < PARALLEL_TOLERANCE;
        }

        // STEP 3: Create dimension chains with ENHANCED collinearity logic and semantic grouping
        private Dictionary<XYZ, List<ProjectedItem>> CreateCombinedDimensionGroups(Dictionary<XYZ, List<ProjectedItem>> parallelGroups)
        {
            var dimensionGroups = new Dictionary<XYZ, List<ProjectedItem>>();

            foreach (var parallelGroup in parallelGroups)
            {
                XYZ direction = parallelGroup.Key;
                List<ProjectedItem> allParallelItems = parallelGroup.Value;

                // Calculate position along direction for sorting
                foreach (var item in allParallelItems)
                {
                    item.PositionAlongDirection = item.ProjectedPoint.DotProduct(direction);
                }

                // Check if we have any selected elements in this parallel group
                var selectedElements = allParallelItems.Where(i =>
                    (i.ItemType == "Element" || i.ItemType == "PerpendicularElement" ||
                     i.ItemType == "CurtainWall" || i.ItemType == "Mullion" ||
                     i.ItemType == "CurtainGridLine") && i.IsSelected).ToList();

                if (selectedElements.Count == 0) continue;

                // Find collinear groups with adaptive tolerances
                List<List<ProjectedItem>> trueCollinearGroups = FindTrueCollinearGroups(allParallelItems, direction);

                // Select one representative from each collinear group
                List<ProjectedItem> representatives = new List<ProjectedItem>();
                foreach (var group in trueCollinearGroups)
                {
                    ProjectedItem rep = SelectBestRepresentative(group);
                    representatives.Add(rep);
                }

                // Create final dimension chain
                if (representatives.Count >= 2)
                {
                    representatives.Sort((a, b) => a.PositionAlongDirection.CompareTo(b.PositionAlongDirection));

                    // ADD THIS DEBUG CALL before adding to dimension groups:
                    LogToFile($"\nDebugging representatives for direction ({direction.X:F3}, {direction.Y:F3}, {direction.Z:F3}):");
                    DebugProjectedItemReferences(representatives);

                    dimensionGroups[direction] = representatives;
                }
            }

            return dimensionGroups;
        }

        private List<List<ProjectedItem>> FindTrueCollinearGroups(List<ProjectedItem> parallelItems, XYZ groupDirection)
        {
            var trueCollinearGroups = new List<List<ProjectedItem>>();

            foreach (var item in parallelItems)
            {
                bool addedToExistingGroup = false;
                double itemTolerance = _settings.GetToleranceForElement(item.Element);

                foreach (var group in trueCollinearGroups)
                {
                    // Use adaptive tolerance based on element types in the group
                    double groupTolerance = Math.Max(itemTolerance, _settings.GetToleranceForElement(group[0].Element));

                    if (LiesOnSameInfiniteLine(item, group[0], groupDirection, groupTolerance))
                    {
                        group.Add(item);
                        addedToExistingGroup = true;
                        break;
                    }
                }

                if (!addedToExistingGroup)
                {
                    trueCollinearGroups.Add(new List<ProjectedItem> { item });
                }
            }

            return trueCollinearGroups;
        }

        private bool LiesOnSameInfiniteLine(ProjectedItem item1, ProjectedItem item2, XYZ groupDirection, double tolerance)
        {
            XYZ perpDirection = new XYZ(-groupDirection.Y, groupDirection.X, 0);
            if (perpDirection.GetLength() < 0.001)
            {
                perpDirection = new XYZ(0, 1, 0);
            }
            perpDirection = perpDirection.Normalize();

            XYZ pointDifference = item2.ProjectedPoint - item1.ProjectedPoint;
            double perpendicularDistance = Math.Abs(pointDifference.DotProduct(perpDirection));

            return perpendicularDistance <= tolerance;
        }

        private ProjectedItem SelectBestRepresentative(List<ProjectedItem> group)
        {
            if (group.Count == 1) return group[0];

            // Enhanced priority: 
            // 1. Selected Elements (linear first)
            // 2. Curtain Wall Grid Lines (most precise)
            // 3. Grids 
            // 4. Levels 
            // 5. Curtain Wall Elements 
            // 6. Other Elements

            var selectedElements = group.Where(i =>
                (i.ItemType == "Element" || i.ItemType == "PerpendicularElement" ||
                 i.ItemType == "CurtainWall" || i.ItemType == "Mullion" || i.ItemType == "CurtainGridLine")
                && i.IsSelected).ToList();

            if (selectedElements.Count > 0)
            {
                // Among selected elements, prefer curtain grid lines for their precision
                var curtainGridLines = selectedElements.Where(i => i.ItemType == "CurtainGridLine").ToList();
                if (curtainGridLines.Count > 0)
                {
                    // Prefer linear grid lines over point grid lines
                    var linearGridLines = curtainGridLines.Where(i => !i.IsPointElement).ToList();
                    if (linearGridLines.Count > 0)
                    {
                        return linearGridLines.First();
                    }
                    return curtainGridLines.First();
                }

                // Prefer linear selected elements over point elements
                var linearSelected = selectedElements.Where(i => !i.IsPointElement).ToList();
                if (linearSelected.Count > 0)
                {
                    // Among linear elements, prefer curtain walls and mullions for their precision
                    var selectedCurtainElements = linearSelected.Where(i => i.IsCurtainWallElement).ToList();
                    if (selectedCurtainElements.Count > 0)
                    {
                        return selectedCurtainElements.First();
                    }
                    return linearSelected.First();
                }
                return selectedElements.First();
            }

            // If no selected elements, check for curtain grid lines (they're very precise)
            var nonSelectedGridLines = group.Where(i => i.ItemType == "CurtainGridLine").ToList();
            if (nonSelectedGridLines.Count > 0)
            {
                var linearNonSelectedGridLines = nonSelectedGridLines.Where(i => !i.IsPointElement).ToList();
                if (linearNonSelectedGridLines.Count > 0)
                {
                    return linearNonSelectedGridLines.First();
                }
                return nonSelectedGridLines.First();
            }

            var grids = group.Where(i => i.ItemType == "Grid").ToList();
            if (grids.Count > 0)
            {
                var namedGrid = grids.FirstOrDefault(g => !string.IsNullOrEmpty(g.Element.Name));
                return namedGrid ?? grids.First();
            }

            var levels = group.Where(i => i.ItemType == "Level").ToList();
            if (levels.Count > 0)
            {
                var namedLevel = levels.FirstOrDefault(l => !string.IsNullOrEmpty(l.Element.Name));
                return namedLevel ?? levels.First();
            }

            var nonSelectedCurtainElements = group.Where(i => i.IsCurtainWallElement).ToList();
            if (nonSelectedCurtainElements.Count > 0)
            {
                return nonSelectedCurtainElements.First();
            }

            var linearElements = group.Where(i => !i.IsPointElement).ToList();
            if (linearElements.Count > 0)
            {
                return linearElements.First();
            }

            return group.First();
        }

        // STEP 4: Create dimension chain with proper reference types
        // REPLACE your CreateDimensionChain method with this debug version:
        private Dimension CreateDimensionChain(Document doc, View view, XYZ direction, List<ProjectedItem> items)
        {
            try
            {
                LogToFile($"\n--- CREATING DIMENSION CHAIN ---");
                LogToFile($"Direction: ({direction.X:F3}, {direction.Y:F3}, {direction.Z:F3})");
                LogToFile($"Items count: {items.Count}");

                // Create reference array
                ReferenceArray refArray = new ReferenceArray();
                LogToFile($"Building reference array:");

                for (int i = 0; i < items.Count; i++)
                {
                    var item = items[i];
                    try
                    {
                        LogToFile($"  Processing item {i + 1}: {item.ItemType} - {item.Element?.Name ?? "Unnamed"} (ID: {item.Element?.Id?.IntegerValue})");

                        Reference reference = item.GetReference(_settings.ReferenceType);
                        if (reference != null)
                        {
                            try
                            {
                                refArray.Append(reference);
                                LogToFile($"    ✅ Reference added successfully");
                            }
                            catch (Exception refEx)
                            {
                                LogToFile($"    ❌ Error appending reference: {refEx.Message}");
                            }
                        }
                        else
                        {
                            LogToFile($"    ❌ GetReference returned null");
                        }
                    }
                    catch (Exception itemEx)
                    {
                        LogToFile($"    ❌ Error processing item: {itemEx.Message}");
                    }
                }

                if (refArray.Size < 2)
                {
                    LogToFile($"❌ Insufficient valid references ({refArray.Size}) for dimension chain");
                    return null;
                }

                LogToFile($"✅ Reference array built with {refArray.Size} references");

                // Calculate dimension line position
                LogToFile($"Calculating dimension line...");

                XYZ perpDirection = new XYZ(-direction.Y, direction.X, 0);
                if (perpDirection.GetLength() < 0.001)
                {
                    perpDirection = new XYZ(0, 1, 0);
                }
                perpDirection = perpDirection.Normalize();

                LogToFile($"Perpendicular direction: ({perpDirection.X:F3}, {perpDirection.Y:F3}, {perpDirection.Z:F3})");

                XYZ groupCenter = XYZ.Zero;
                foreach (var item in items)
                {
                    groupCenter += item.ProjectedPoint;
                }
                groupCenter = groupCenter / items.Count;
                LogToFile($"Group center: ({groupCenter.X:F3}, {groupCenter.Y:F3}, {groupCenter.Z:F3})");

                XYZ dimLinePoint2D = groupCenter + perpDirection * DIMENSION_OFFSET;
                XYZ dimLinePoint3D = ConvertViewPointTo3D(dimLinePoint2D, view);
                XYZ direction3D = ConvertViewDirectionTo3D(direction, view);

                if (direction3D.GetLength() < 0.001)
                {
                    direction3D = new XYZ(1, 0, 0);
                }
                else
                {
                    direction3D = direction3D.Normalize();
                }

                LogToFile($"Dimension line point 3D: ({dimLinePoint3D.X:F3}, {dimLinePoint3D.Y:F3}, {dimLinePoint3D.Z:F3})");
                LogToFile($"Direction 3D: ({direction3D.X:F3}, {direction3D.Y:F3}, {direction3D.Z:F3})");

                double minProj = items.Min(item => item.PositionAlongDirection);
                double maxProj = items.Max(item => item.PositionAlongDirection);
                LogToFile($"Projection range: {minProj:F3} to {maxProj:F3}");

                double span = Math.Abs(maxProj - minProj);
                if (span < 1.0)
                {
                    double center = (minProj + maxProj) / 2;
                    minProj = center - 0.5;
                    maxProj = center + 0.5;
                    LogToFile($"Adjusted projection range: {minProj:F3} to {maxProj:F3}");
                }

                XYZ dimStart3D = dimLinePoint3D + direction3D * (minProj - 1.97);
                XYZ dimEnd3D = dimLinePoint3D + direction3D * (maxProj + 1.97);

                LogToFile($"Dimension start: ({dimStart3D.X:F3}, {dimStart3D.Y:F3}, {dimStart3D.Z:F3})");
                LogToFile($"Dimension end: ({dimEnd3D.X:F3}, {dimEnd3D.Y:F3}, {dimEnd3D.Z:F3})");

                // Validate dimension line
                double dimLineLength = dimStart3D.DistanceTo(dimEnd3D);
                LogToFile($"Dimension line length: {dimLineLength:F3}");

                if (dimLineLength < 0.1) // Less than about 1 inch
                {
                    LogToFile($"❌ Dimension line too short: {dimLineLength:F3}");
                    return null;
                }

                Line dimensionLine = Line.CreateBound(dimStart3D, dimEnd3D);
                LogToFile($"✅ Dimension line created successfully");

                // Additional validation
                LogToFile($"View information:");
                LogToFile($"  View name: {view.Name}");
                LogToFile($"  View type: {view.ViewType}");
                LogToFile($"  View is template: {view.IsTemplate}");
                LogToFile($"  View scale: {view.Scale}");

                // Check if view is locked or has issues
                try
                {
                    var cropParam = view.get_Parameter(BuiltInParameter.VIEWER_CROP_REGION_VISIBLE);
                    if (cropParam != null)
                    {
                        LogToFile($"  View crop region visible: {cropParam.AsInteger() == 1}");
                    }
                }
                catch
                {
                    LogToFile($"  Could not check view crop status");
                }

                LogToFile($"Attempting to create dimension...");
                LogToFile($"  Reference array size: {refArray.Size}");
                LogToFile($"  Dimension line length: {dimensionLine.Length:F3}");

                // Try to create the dimension
                Dimension dimension = null;
                try
                {
                    dimension = doc.Create.NewDimension(view, dimensionLine, refArray);

                    if (dimension != null)
                    {
                        LogToFile($"✅ Dimension created successfully! ID: {dimension.Id.IntegerValue}");

                        // Additional dimension info
                        LogToFile($"  Dimension segments: {dimension.Segments.Size}");
                        LogToFile($"  Dimension curve length: {dimension.Curve.Length:F3}");

                        // Try to get dimension value
                        try
                        {
                            string dimValue = dimension.ValueString;
                            LogToFile($"  Dimension value: {dimValue}");
                        }
                        catch (Exception valueEx)
                        {
                            LogToFile($"  Could not get dimension value: {valueEx.Message}");
                        }
                    }
                    else
                    {
                        LogToFile($"❌ doc.Create.NewDimension returned null - this usually means:");
                        LogToFile($"  1. Invalid references (elements not dimensionable)");
                        LogToFile($"  2. References are collinear with dimension line");
                        LogToFile($"  3. View doesn't support dimensions");
                        LogToFile($"  4. Transaction not active");
                    }
                }
                catch (Autodesk.Revit.Exceptions.ArgumentException argEx)
                {
                    LogToFile($"❌ ArgumentException creating dimension: {argEx.Message}");
                    LogToFile($"  This usually means invalid references or geometry");
                }
                catch (Autodesk.Revit.Exceptions.InvalidOperationException invOpEx)
                {
                    LogToFile($"❌ InvalidOperationException creating dimension: {invOpEx.Message}");
                    LogToFile($"  This usually means transaction issues or view problems");
                }
                catch (Exception genEx)
                {
                    LogToFile($"❌ General exception creating dimension: {genEx.Message}");
                    LogToFile($"  Exception type: {genEx.GetType().Name}");
                    LogToFile($"  Stack trace: {genEx.StackTrace}");
                }

                return dimension;
            }
            catch (Exception ex)
            {
                LogToFile($"❌ Error in CreateDimensionChain: {ex.Message}");
                LogToFile($"Exception type: {ex.GetType().Name}");
                LogToFile($"Stack trace: {ex.StackTrace}");
                return null;
            }
        }

        private void DebugProjectedItemReferences(List<ProjectedItem> items)
        {
            LogToFile($"\n--- DEBUGGING PROJECTED ITEM REFERENCES ---");

            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                LogToFile($"Item {i + 1}: {item.ItemType} - ID: {item.Element?.Id?.IntegerValue}");

                try
                {
                    // Test getting different reference types
                    var centerlineRef = item.CenterlineReference;
                    var geometricRef = item.GeometricReference;
                    var settingsRef = item.GetReference(_settings.ReferenceType);

                    LogToFile($"  Centerline ref: {(centerlineRef != null ? "✅ Valid" : "❌ Null")}");
                    LogToFile($"  Geometric ref: {(geometricRef != null ? "✅ Valid" : "❌ Null")}");
                    LogToFile($"  Settings ref ({_settings.ReferenceType}): {(settingsRef != null ? "✅ Valid" : "❌ Null")}");

                    // Test if element is valid for dimensioning
                    if (item.Element != null)
                    {
                        LogToFile($"  Element valid: ✅");
                        LogToFile($"  Element category: {item.Element.Category?.Name ?? "No Category"}");

                        // For curtain grid lines, try to get additional info
                        if (item.Element is CurtainGridLine gridLine)
                        {
                            LogToFile($"  CurtainGridLine details:");
                            try
                            {
                                var curve = gridLine.FullCurve;
                                LogToFile($"    Has curve: {(curve != null ? "✅" : "❌")}");
                                if (curve != null)
                                {
                                    LogToFile($"    Curve type: {curve.GetType().Name}");
                                    LogToFile($"    Curve length: {curve.Length:F3}");
                                }
                            }
                            catch (Exception curveEx)
                            {
                                LogToFile($"    Error getting curve: {curveEx.Message}");
                            }
                        }
                    }
                    else
                    {
                        LogToFile($"  Element: ❌ Null");
                    }
                }
                catch (Exception refEx)
                {
                    LogToFile($"  Error checking references: {refEx.Message}");
                }

                LogToFile($""); // Empty line for readability
            }
        }

        private XYZ ConvertViewPointTo3D(XYZ point2D, View view)
        {
            switch (view.ViewType)
            {
                case ViewType.FloorPlan:
                case ViewType.CeilingPlan:
                case ViewType.AreaPlan:
                    return new XYZ(point2D.X, point2D.Y, 0);

                case ViewType.Section:
                case ViewType.Elevation:
                    XYZ viewOrigin = view.Origin;
                    XYZ viewRight = view.RightDirection.Normalize();
                    XYZ viewUp = view.UpDirection.Normalize();
                    return viewOrigin + viewRight * point2D.X + viewUp * point2D.Y;

                default:
                    return new XYZ(point2D.X, point2D.Y, 0);
            }
        }

        private XYZ ConvertViewDirectionTo3D(XYZ direction2D, View view)
        {
            switch (view.ViewType)
            {
                case ViewType.FloorPlan:
                case ViewType.CeilingPlan:
                case ViewType.AreaPlan:
                    return new XYZ(direction2D.X, direction2D.Y, 0).Normalize();

                case ViewType.Section:
                case ViewType.Elevation:
                    XYZ viewRight = view.RightDirection.Normalize();
                    XYZ viewUp = view.UpDirection.Normalize();
                    return (viewRight * direction2D.X + viewUp * direction2D.Y).Normalize();

                default:
                    return new XYZ(direction2D.X, direction2D.Y, 0).Normalize();
            }
        }

        // STEP 5: Move dimensions left with enhanced positioning
        private void MoveDimensionsLeft(Document doc, List<Dimension> dimensions, double moveDistance)
        {
            foreach (Dimension dimension in dimensions)
            {
                try
                {
                    Line dimensionLine = dimension.Curve as Line;
                    if (dimensionLine == null) continue;

                    XYZ dimensionDirection = dimensionLine.Direction.Normalize();
                    XYZ leftDirection = new XYZ(-dimensionDirection.Y, dimensionDirection.X, 0);

                    if (leftDirection.GetLength() < 0.1)
                        leftDirection = new XYZ(-1, 0, 0);

                    leftDirection = leftDirection.Normalize();
                    XYZ translationVector = leftDirection * moveDistance;

                    ElementTransformUtils.MoveElement(doc, dimension.Id, translationVector);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error moving dimension: {ex.Message}");
                }
            }
        }
    }
}