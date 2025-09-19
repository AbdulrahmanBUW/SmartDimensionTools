using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;

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
                // Get all curtain walls in the view
                var curtainWalls = GetCurtainWallsInView(doc, view);

                foreach (var curtainWall in curtainWalls)
                {
                    try
                    {
                        // Process the curtain wall itself if selected
                        if (selectedIds.Contains(curtainWall.Id) && settings.IncludeCurtainWalls)
                        {
                            var wallItem = ProcessCurtainWall(curtainWall, view, settings);
                            if (wallItem != null)
                            {
                                curtainWallItems.Add(wallItem);
                            }
                        }

                        // Process mullions if enabled
                        if (settings.IncludeMullions)
                        {
                            var mullionItems = ProcessCurtainWallMullions(curtainWall, view, selectedIds, settings);
                            curtainWallItems.AddRange(mullionItems);
                        }

                        // Process panels if needed
                        var panelItems = ProcessCurtainWallPanels(curtainWall, view, selectedIds, settings);
                        curtainWallItems.AddRange(panelItems);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error processing curtain wall {curtainWall.Id}: {ex.Message}");
                    }
                }

                System.Diagnostics.Debug.WriteLine($"Found {curtainWallItems.Count} curtain wall elements");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ProcessCurtainWallElements: {ex.Message}");
            }

            return curtainWallItems;
        }

        /// <summary>
        /// Gets all curtain walls visible in the view
        /// </summary>
        private static List<Wall> GetCurtainWallsInView(Document doc, View view)
        {
            var curtainWalls = new List<Wall>();

            try
            {
                var collector = new FilteredElementCollector(doc, view.Id)
                    .OfClass(typeof(Wall));

                foreach (Wall wall in collector)
                {
                    if (IsCurtainWall(wall))
                    {
                        curtainWalls.Add(wall);
                    }
                }

                // Also get all curtain walls from project (might not be visible in current view filter)
                var allWallsCollector = new FilteredElementCollector(doc)
                    .OfClass(typeof(Wall));

                foreach (Wall wall in allWallsCollector)
                {
                    if (IsCurtainWall(wall) && !curtainWalls.Any(cw => cw.Id == wall.Id))
                    {
                        // Check if wall intersects with view
                        if (IsWallVisibleInView(wall, view))
                        {
                            curtainWalls.Add(wall);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting curtain walls: {ex.Message}");
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
                return wall.WallType.Kind == WallKind.Curtain;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Checks if wall is visible in the given view
        /// </summary>
        private static bool IsWallVisibleInView(Wall wall, View view)
        {
            try
            {
                // Get wall location curve
                if (!(wall.Location is LocationCurve locationCurve)) return false;

                Curve wallCurve = locationCurve.Curve;

                // For plan views, check Z-coordinate range
                if (view.ViewType == ViewType.FloorPlan ||
                    view.ViewType == ViewType.CeilingPlan ||
                    view.ViewType == ViewType.AreaPlan)
                {
                    var level = view.GenLevel;
                    if (level != null)
                    {
                        double viewElevation = level.Elevation;
                        double wallBottom = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)?.AsDouble() ?? 0;
                        double wallTop = wallBottom + wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)?.AsDouble() ?? 10;

                        return viewElevation >= wallBottom && viewElevation <= wallTop;
                    }
                }

                // For sections/elevations, use geometric intersection
                // (This is a simplified check - could be enhanced)
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error checking wall visibility: {ex.Message}");
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
        /// Process curtain wall mullions
        /// </summary>
        private static List<ProjectedItem> ProcessCurtainWallMullions(Wall curtainWall, View view,
            ICollection<ElementId> selectedIds, DimensionSettings settings)
        {
            var mullionItems = new List<ProjectedItem>();

            try
            {
                // Get curtain grid
                CurtainGrid curtainGrid = curtainWall.CurtainGrid;
                if (curtainGrid == null) return mullionItems;

                // Process U-direction mullions (typically vertical)
                foreach (ElementId mullionId in curtainGrid.GetUGridLineIds())
                {
                    try
                    {
                        var mullion = curtainWall.Document.GetElement(mullionId) as Mullion;
                        if (mullion != null)
                        {
                            // Check if this mullion should be included
                            bool includeThis = selectedIds.Count == 0 || // Include all if nothing selected
                                             selectedIds.Contains(mullionId) || // Specifically selected
                                             selectedIds.Contains(curtainWall.Id); // Parent wall selected

                            if (includeThis)
                            {
                                var mullionItem = ProcessMullion(mullion, curtainWall, view, settings, "U");
                                if (mullionItem != null)
                                {
                                    mullionItems.Add(mullionItem);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error processing U-mullion {mullionId}: {ex.Message}");
                    }
                }

                // Process V-direction mullions (typically horizontal)
                foreach (ElementId mullionId in curtainGrid.GetVGridLineIds())
                {
                    try
                    {
                        var mullion = curtainWall.Document.GetElement(mullionId) as Mullion;
                        if (mullion != null)
                        {
                            bool includeThis = selectedIds.Count == 0 ||
                                             selectedIds.Contains(mullionId) ||
                                             selectedIds.Contains(curtainWall.Id);

                            if (includeThis)
                            {
                                var mullionItem = ProcessMullion(mullion, curtainWall, view, settings, "V");
                                if (mullionItem != null)
                                {
                                    mullionItems.Add(mullionItem);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error processing V-mullion {mullionId}: {ex.Message}");
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
        /// Process a single mullion
        /// </summary>
        private static ProjectedItem ProcessMullion(Mullion mullion, Wall parentWall, View view,
            DimensionSettings settings, string direction)
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
                        ReferenceType = DimensionReferenceType.Centerline, // Mullions always use centerline
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

        /// <summary>
        /// Process curtain wall panels (if needed)
        /// </summary>
        private static List<ProjectedItem> ProcessCurtainWallPanels(Wall curtainWall, View view,
            ICollection<ElementId> selectedIds, DimensionSettings settings)
        {
            var panelItems = new List<ProjectedItem>();

            // For now, we typically don't dimension to panels, but this could be extended
            // if users need to dimension to panel edges

            return panelItems;
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