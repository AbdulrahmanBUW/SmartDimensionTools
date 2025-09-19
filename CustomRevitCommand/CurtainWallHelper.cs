using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CustomRevitCommand
{
    /// <summary>
    /// Helper class for processing curtain walls and their components
    /// </summary>
    public static class CurtainWallHelper
    {
        /// <summary>
        /// Processes curtain wall elements and returns projected items
        /// </summary>
        public static List<ProjectedItem> ProcessCurtainWallElements(Document doc, View view,
            ICollection<ElementId> selectedIds, DimensionSettings settings)
        {
            var curtainWallItems = new List<ProjectedItem>();

            try
            {
                // Get all curtain walls in the project (not just in view)
                var allCurtainWalls = GetAllCurtainWalls(doc);
                System.Diagnostics.Debug.WriteLine($"Found {allCurtainWalls.Count} curtain walls in project");

                foreach (var curtainWall in allCurtainWalls)
                {
                    try
                    {
                        // Check if this curtain wall is relevant to the current view
                        if (!IsCurtainWallRelevantToView(curtainWall, view))
                            continue;

                        // Process the curtain wall itself if selected and settings allow
                        if (selectedIds.Contains(curtainWall.Id) && settings.IncludeCurtainWalls)
                        {
                            var wallItem = ProcessCurtainWall(curtainWall, view, settings);
                            if (wallItem != null)
                            {
                                curtainWallItems.Add(wallItem);
                                System.Diagnostics.Debug.WriteLine($"Added curtain wall: {curtainWall.Id}");
                            }
                        }

                        // Process curtain wall grids if enabled
                        if (settings.IncludeMullions)
                        {
                            var gridItems = ProcessCurtainWallGrids(curtainWall, view, selectedIds, settings);
                            curtainWallItems.AddRange(gridItems);
                            System.Diagnostics.Debug.WriteLine($"Added {gridItems.Count} grid items from wall {curtainWall.Id}");
                        }

                        // Process actual mullions if they exist
                        var mullionItems = ProcessCurtainWallMullions(curtainWall, view, selectedIds, settings);
                        curtainWallItems.AddRange(mullionItems);
                        System.Diagnostics.Debug.WriteLine($"Added {mullionItems.Count} mullion items from wall {curtainWall.Id}");

                        // Process panels if needed (usually not for dimensioning)
                        var panelItems = ProcessCurtainWallPanels(curtainWall, view, selectedIds, settings);
                        curtainWallItems.AddRange(panelItems);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error processing curtain wall {curtainWall.Id}: {ex.Message}");
                    }
                }

                System.Diagnostics.Debug.WriteLine($"Total curtain wall elements found: {curtainWallItems.Count}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ProcessCurtainWallElements: {ex.Message}");
            }

            return curtainWallItems;
        }

        /// <summary>
        /// Gets all curtain walls from the entire project
        /// </summary>
        private static List<Wall> GetAllCurtainWalls(Document doc)
        {
            var curtainWalls = new List<Wall>();

            try
            {
                var allWalls = new FilteredElementCollector(doc)
                    .OfClass(typeof(Wall))
                    .Cast<Wall>();

                foreach (var wall in allWalls)
                {
                    if (IsCurtainWall(wall))
                    {
                        curtainWalls.Add(wall);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting all curtain walls: {ex.Message}");
            }

            return curtainWalls;
        }

        /// <summary>
        /// Checks if a wall is a curtain wall
        /// </summary>
        public static bool IsCurtainWall(Wall wall)
        {
            try
            {
                return wall != null && wall.WallType?.Kind == WallKind.Curtain;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Checks if curtain wall is relevant to the current view
        /// </summary>
        private static bool IsCurtainWallRelevantToView(Wall curtainWall, View view)
        {
            try
            {
                // Get wall location curve
                if (!(curtainWall.Location is LocationCurve locationCurve))
                    return false;

                var wallCurve = locationCurve.Curve;
                var wallStart = wallCurve.GetEndPoint(0);
                var wallEnd = wallCurve.GetEndPoint(1);

                switch (view.ViewType)
                {
                    case ViewType.FloorPlan:
                    case ViewType.CeilingPlan:
                    case ViewType.AreaPlan:
                        // Check if wall intersects with view's level range
                        var viewLevel = view.GenLevel;
                        if (viewLevel != null)
                        {
                            double viewElevation = viewLevel.Elevation;
                            double wallBottom = GetWallBottomElevation(curtainWall);
                            double wallTop = GetWallTopElevation(curtainWall);

                            // Include if view elevation is within wall height range
                            return viewElevation >= wallBottom - 1.0 && viewElevation <= wallTop + 1.0;
                        }
                        return true; // Default to include if can't determine

                    case ViewType.Section:
                    case ViewType.Elevation:
                        // For sections/elevations, include all walls
                        return true;

                    default:
                        return false;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error checking wall relevance: {ex.Message}");
                return true; // Default to include on error
            }
        }

        /// <summary>
        /// Process curtain wall grids - SIMPLIFIED VERSION
        /// </summary>
        private static List<ProjectedItem> ProcessCurtainWallGrids(Wall curtainWall, View view,
            ICollection<ElementId> selectedIds, DimensionSettings settings)
        {
            var gridItems = new List<ProjectedItem>();

            try
            {
                CurtainGrid curtainGrid = curtainWall.CurtainGrid;
                if (curtainGrid == null)
                {
                    System.Diagnostics.Debug.WriteLine($"No curtain grid found for wall {curtainWall.Id}");
                    return gridItems;
                }

                System.Diagnostics.Debug.WriteLine($"Processing curtain grid for wall {curtainWall.Id}");
                System.Diagnostics.Debug.WriteLine($"Note: CurtainGridLine elements are not directly dimensionable in Revit.");
                System.Diagnostics.Debug.WriteLine($"Dimensioning to curtain wall centerline instead.");

                // Instead of trying to process grid lines, create reference points from curtain wall geometry
                var referenceItems = CreateCurtainWallReferencePoints(curtainWall, view, settings);
                gridItems.AddRange(referenceItems);

                System.Diagnostics.Debug.WriteLine($"Created {referenceItems.Count} curtain wall reference points");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error processing curtain wall grids: {ex.Message}");
            }

            return gridItems;
        }

        /// <summary>
        /// Create dimensionable reference points from curtain wall geometry
        /// </summary>
        private static List<ProjectedItem> CreateCurtainWallReferencePoints(Wall curtainWall, View view, DimensionSettings settings)
        {
            var referenceItems = new List<ProjectedItem>();

            try
            {
                if (!(curtainWall.Location is LocationCurve locationCurve)) return referenceItems;

                var wallCurve = locationCurve.Curve;
                var wallStart = wallCurve.GetEndPoint(0);
                var wallEnd = wallCurve.GetEndPoint(1);

                // Project to view
                var start2D = ProjectPointToView(wallStart, view);
                var end2D = ProjectPointToView(wallEnd, view);
                var direction2D = (end2D - start2D);

                if (direction2D.GetLength() < 0.001) return referenceItems;

                direction2D = direction2D.Normalize();

                // For now, we'll just create one reference item for the curtain wall centerline
                // This avoids the issue with non-dimensionable grid lines
                var centerItem = new ProjectedItem
                {
                    Element = curtainWall,
                    GeometricReference = new Reference(curtainWall),
                    CenterlineReference = new Reference(curtainWall),
                    ProjectedDirection = direction2D,
                    ProjectedPoint = (start2D + end2D) / 2, // Wall center
                    ItemType = "CurtainWallCenter",
                    IsSelected = true,
                    IsPointElement = false,
                    IsCurtainWallElement = true,
                    IsMullion = false,
                    ReferenceType = DimensionReferenceType.Centerline,
                    ElementWidth = GetWallWidth(curtainWall)
                };

                referenceItems.Add(centerItem);
                System.Diagnostics.Debug.WriteLine($"Created curtain wall center reference point");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating curtain wall reference points: {ex.Message}");
            }

            return referenceItems;
        }

        /// <summary>
        /// Process curtain wall mullions (actual mullion elements)
        /// </summary>
        private static List<ProjectedItem> ProcessCurtainWallMullions(Wall curtainWall, View view,
            ICollection<ElementId> selectedIds, DimensionSettings settings)
        {
            var mullionItems = new List<ProjectedItem>();

            // Only process mullions if the setting specifically includes them
            if (!settings.IncludeMullions) return mullionItems;

            try
            {
                CurtainGrid curtainGrid = curtainWall.CurtainGrid;
                if (curtainGrid == null) return mullionItems;

                // Get all mullions from the curtain grid
                var allMullions = GetAllMullionsFromGrid(curtainGrid, curtainWall.Document);
                System.Diagnostics.Debug.WriteLine($"Found {allMullions.Count} mullions in curtain wall {curtainWall.Id}");

                foreach (var mullion in allMullions)
                {
                    try
                    {
                        bool includeThis = ShouldIncludeGridLine(mullion.Id, curtainWall.Id, selectedIds);

                        if (includeThis)
                        {
                            var mullionItem = ProcessMullion(mullion, curtainWall, view, settings);
                            if (mullionItem != null)
                            {
                                mullionItems.Add(mullionItem);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error processing mullion {mullion.Id}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error processing curtain wall mullions: {ex.Message}");
            }

            return mullionItems;
        }

        /// <summary>
        /// Process curtain wall panels (if needed)
        /// </summary>
        private static List<ProjectedItem> ProcessCurtainWallPanels(Wall curtainWall, View view,
            ICollection<ElementId> selectedIds, DimensionSettings settings)
        {
            var panelItems = new List<ProjectedItem>();
            // For now, we typically don't dimension to panels
            return panelItems;
        }

        /// <summary>
        /// Determines if a grid line should be included based on selection
        /// </summary>
        private static bool ShouldIncludeGridLine(ElementId gridLineId, ElementId parentWallId, ICollection<ElementId> selectedIds)
        {
            // Include if:
            // 1. Nothing is selected (include all)
            // 2. The grid line itself is selected
            // 3. The parent curtain wall is selected

            if (selectedIds.Count == 0) return true;
            if (selectedIds.Contains(gridLineId)) return true;
            if (selectedIds.Contains(parentWallId)) return true;

            return false;
        }

        /// <summary>
        /// Get all mullions from a curtain grid
        /// </summary>
        private static List<Mullion> GetAllMullionsFromGrid(CurtainGrid curtainGrid, Document doc)
        {
            var mullions = new List<Mullion>();

            try
            {
                // Get mullions directly from the document
                var mullionCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(Mullion))
                    .Cast<Mullion>();

                foreach (var mullion in mullionCollector)
                {
                    // Check if this mullion belongs to this curtain grid
                    if (IsMullionInCurtainGrid(mullion, curtainGrid) && !mullions.Any(m => m.Id == mullion.Id))
                    {
                        mullions.Add(mullion);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting mullions from grid: {ex.Message}");
            }

            return mullions;
        }

        /// <summary>
        /// Check if a mullion belongs to a specific curtain grid
        /// </summary>
        private static bool IsMullionInCurtainGrid(Mullion mullion, CurtainGrid curtainGrid)
        {
            try
            {
                // Simplified approach - in a full implementation, you'd need more sophisticated analysis
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Process the curtain wall itself
        /// </summary>
        private static ProjectedItem ProcessCurtainWall(Wall curtainWall, View view, DimensionSettings settings)
        {
            try
            {
                if (!(curtainWall.Location is LocationCurve locationCurve)) return null;

                Curve curve = locationCurve.Curve;
                XYZ start3D = curve.GetEndPoint(0);
                XYZ end3D = curve.GetEndPoint(1);

                // Project to view surface
                XYZ start2D = ProjectPointToView(start3D, view);
                XYZ end2D = ProjectPointToView(end3D, view);

                XYZ direction2D = (end2D - start2D);
                if (direction2D.GetLength() < 0.001) return null;

                direction2D = direction2D.Normalize();
                XYZ center2D = (start2D + end2D) / 2;

                // Get wall thickness for face calculations
                double wallWidth = GetWallWidth(curtainWall);

                // Get references
                var references = GetWallReferences(curtainWall, settings.ReferenceType);

                var projectedItem = new ProjectedItem
                {
                    Element = curtainWall,
                    GeometricReference = references.Main,
                    CenterlineReference = references.Centerline,
                    ExteriorFaceReference = references.ExteriorFace,
                    InteriorFaceReference = references.InteriorFace,
                    ProjectedDirection = direction2D,
                    ProjectedPoint = center2D,
                    ItemType = "CurtainWall",
                    IsSelected = true,
                    IsPointElement = false,
                    IsCurtainWallElement = true,
                    IsMullion = false,
                    ReferenceType = settings.ReferenceType,
                    ElementWidth = wallWidth
                };

                return projectedItem;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error processing curtain wall: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Process a single mullion
        /// </summary>
        private static ProjectedItem ProcessMullion(Mullion mullion, Wall parentWall, View view,
            DimensionSettings settings)
        {
            try
            {
                // Get mullion geometry
                var mullionGeometry = GetMullionGeometry(mullion, view);
                if (mullionGeometry == null) return null;

                XYZ center3D = mullionGeometry.Center;
                XYZ direction3D = mullionGeometry.Direction;

                // Project to view
                XYZ center2D = ProjectPointToView(center3D, view);
                XYZ direction2D = ProjectDirectionToView(direction3D, view);

                if (direction2D.GetLength() < 0.001)
                {
                    // This is a perpendicular mullion - treat as point element
                    var pointItem = new ProjectedItem
                    {
                        Element = mullion,
                        GeometricReference = new Reference(mullion),
                        CenterlineReference = new Reference(mullion),
                        ProjectedDirection = null,
                        ProjectedPoint = center2D,
                        ItemType = "Mullion",
                        IsSelected = true,
                        IsPointElement = true,
                        IsCurtainWallElement = true,
                        IsMullion = true,
                        ParentWallId = parentWall.Id,
                        ReferenceType = DimensionReferenceType.Centerline,
                        ElementWidth = GetMullionWidth(mullion)
                    };

                    return pointItem;
                }

                direction2D = direction2D.Normalize();

                var projectedItem = new ProjectedItem
                {
                    Element = mullion,
                    GeometricReference = new Reference(mullion),
                    CenterlineReference = new Reference(mullion),
                    ProjectedDirection = direction2D,
                    ProjectedPoint = center2D,
                    ItemType = "Mullion",
                    IsSelected = true,
                    IsPointElement = false,
                    IsCurtainWallElement = true,
                    IsMullion = true,
                    ParentWallId = parentWall.Id,
                    ReferenceType = DimensionReferenceType.Centerline,
                    ElementWidth = GetMullionWidth(mullion)
                };

                return projectedItem;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error processing mullion {mullion.Id}: {ex.Message}");
                return null;
            }
        }

        #region Helper Methods

        private static double GetWallBottomElevation(Wall wall)
        {
            try
            {
                var baseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)?.AsDouble() ?? 0;
                var level = wall.Document.GetElement(wall.LevelId) as Level;
                var levelElevation = level?.Elevation ?? 0;
                return levelElevation + baseOffset;
            }
            catch
            {
                return 0;
            }
        }

        private static double GetWallTopElevation(Wall wall)
        {
            try
            {
                var baseElevation = GetWallBottomElevation(wall);
                var height = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)?.AsDouble() ?? 10;
                return baseElevation + height;
            }
            catch
            {
                return 10; // Default height
            }
        }

        /// <summary>
        /// Get mullion geometry information
        /// </summary>
        private static MullionGeometry GetMullionGeometry(Mullion mullion, View view)
        {
            try
            {
                // Get mullion location curve
                if (!(mullion.Location is LocationCurve locationCurve)) return null;

                Curve curve = locationCurve.Curve;
                XYZ start = curve.GetEndPoint(0);
                XYZ end = curve.GetEndPoint(1);
                XYZ center = (start + end) / 2;
                XYZ direction = (end - start).Normalize();

                return new MullionGeometry
                {
                    Center = center,
                    Direction = direction,
                    Start = start,
                    End = end
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting mullion geometry: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Get wall references for different reference types
        /// </summary>
        private static WallReferences GetWallReferences(Wall wall, DimensionReferenceType referenceType)
        {
            try
            {
                var references = new WallReferences();

                // Get wall faces
                references.Main = new Reference(wall);
                references.Centerline = new Reference(wall);

                // Try to get face references
                var faces = GetWallFaces(wall);
                if (faces != null)
                {
                    references.ExteriorFace = faces.ExteriorFace;
                    references.InteriorFace = faces.InteriorFace;
                }

                return references;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting wall references: {ex.Message}");
                return new WallReferences { Main = new Reference(wall), Centerline = new Reference(wall) };
            }
        }

        /// <summary>
        /// Get wall face references
        /// </summary>
        private static WallFaces GetWallFaces(Wall wall)
        {
            try
            {
                // This is complex in Revit - simplified approach
                // In a full implementation, you'd analyze the wall geometry to find face references

                return new WallFaces
                {
                    ExteriorFace = new Reference(wall),
                    InteriorFace = new Reference(wall)
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting wall faces: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Get wall width/thickness
        /// </summary>
        private static double GetWallWidth(Wall wall)
        {
            try
            {
                var wallType = wall.WallType;
                if (wallType != null)
                {
                    var widthParam = wallType.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                    if (widthParam != null)
                    {
                        return widthParam.AsDouble();
                    }
                }

                // Fallback - try to get from compound structure
                var compoundStructure = wallType?.GetCompoundStructure();
                if (compoundStructure != null)
                {
                    return compoundStructure.GetWidth();
                }

                return 0.5; // Default 6 inches
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting wall width: {ex.Message}");
                return 0.5;
            }
        }

        /// <summary>
        /// Get mullion width
        /// </summary>
        private static double GetMullionWidth(Mullion mullion)
        {
            try
            {
                var mullionType = mullion.Document.GetElement(mullion.GetTypeId()) as MullionType;
                if (mullionType != null)
                {
                    // Try different possible parameter names for mullion width
                    Parameter widthParam = null;

                    // Look through all parameters to find width-related ones
                    foreach (Parameter param in mullionType.Parameters)
                    {
                        string paramName = param.Definition.Name.ToLower();
                        if (paramName.Contains("width") || paramName.Contains("depth"))
                        {
                            widthParam = param;
                            break;
                        }
                    }

                    if (widthParam != null && widthParam.HasValue)
                    {
                        return widthParam.AsDouble();
                    }
                }

                return 0.25; // Default 3 inches
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting mullion width: {ex.Message}");
                return 0.25;
            }
        }

        /// <summary>
        /// Project 3D point to view surface
        /// </summary>
        private static XYZ ProjectPointToView(XYZ point3D, View view)
        {
            switch (view.ViewType)
            {
                case ViewType.FloorPlan:
                case ViewType.CeilingPlan:
                case ViewType.AreaPlan:
                    return new XYZ(point3D.X, point3D.Y, 0);

                case ViewType.Section:
                case ViewType.Elevation:
                    XYZ viewOrigin = view.Origin;
                    XYZ viewRight = view.RightDirection.Normalize();
                    XYZ viewUp = view.UpDirection.Normalize();
                    XYZ relativePoint = point3D - viewOrigin;
                    double x = relativePoint.DotProduct(viewRight);
                    double y = relativePoint.DotProduct(viewUp);
                    return new XYZ(x, y, 0);

                default:
                    return new XYZ(point3D.X, point3D.Y, 0);
            }
        }

        /// <summary>
        /// Project 3D direction to view surface
        /// </summary>
        private static XYZ ProjectDirectionToView(XYZ direction3D, View view)
        {
            switch (view.ViewType)
            {
                case ViewType.FloorPlan:
                case ViewType.CeilingPlan:
                case ViewType.AreaPlan:
                    var dir2D = new XYZ(direction3D.X, direction3D.Y, 0);
                    return dir2D.GetLength() > 0.001 ? dir2D.Normalize() : new XYZ(1, 0, 0);

                case ViewType.Section:
                case ViewType.Elevation:
                    XYZ viewRight = view.RightDirection.Normalize();
                    XYZ viewUp = view.UpDirection.Normalize();
                    double x = direction3D.DotProduct(viewRight);
                    double y = direction3D.DotProduct(viewUp);
                    var projectedDir = new XYZ(x, y, 0);
                    return projectedDir.GetLength() > 0.001 ? projectedDir.Normalize() : new XYZ(1, 0, 0);

                default:
                    var defaultDir = new XYZ(direction3D.X, direction3D.Y, 0);
                    return defaultDir.GetLength() > 0.001 ? defaultDir.Normalize() : new XYZ(1, 0, 0);
            }
        }

        #endregion
    }

    /// <summary>
    /// Helper classes for curtain wall processing
    /// </summary>
    public class MullionGeometry
    {
        public XYZ Center { get; set; }
        public XYZ Direction { get; set; }
        public XYZ Start { get; set; }
        public XYZ End { get; set; }
    }

    public class WallReferences
    {
        public Reference Main { get; set; }
        public Reference Centerline { get; set; }
        public Reference ExteriorFace { get; set; }
        public Reference InteriorFace { get; set; }
    }

    public class WallFaces
    {
        public Reference ExteriorFace { get; set; }
        public Reference InteriorFace { get; set; }
    }
}