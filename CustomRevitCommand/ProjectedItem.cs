using Autodesk.Revit.DB;
using System;

namespace CustomRevitCommand
{
    public class ProjectedItem
    {
        public Element Element { get; set; }
        public Reference GeometricReference { get; set; }
        public XYZ ProjectedDirection { get; set; }
        public XYZ ProjectedPoint { get; set; }
        public double PositionAlongDirection { get; set; }
        public string ItemType { get; set; }
        public bool IsSelected { get; set; }
        public bool IsPointElement { get; set; }

        // New properties for enhanced support
        public DimensionReferenceType ReferenceType { get; set; }
        public Reference ExteriorFaceReference { get; set; }
        public Reference InteriorFaceReference { get; set; }
        public Reference CenterlineReference { get; set; }
        public ElementId ParentWallId { get; set; } // For curtain wall elements
        public bool IsCurtainWallElement { get; set; }
        public bool IsMullion { get; set; }
        public double ElementWidth { get; set; } // For calculating face offsets

        /// <summary>
        /// Gets the appropriate reference based on the specified reference type
        /// </summary>
        public Reference GetReference(DimensionReferenceType referenceType)
        {
            try
            {
                // Special handling for curtain wall grid lines - always use centerline
                if (ItemType == "CurtainGridLine")
                {
                    return CenterlineReference ?? GeometricReference;
                }

                // Special handling for building grids - always use their inherent reference
                if (ItemType == "Grid")
                {
                    return GeometricReference;
                }

                // Special handling for levels - always use their inherent reference
                if (ItemType == "Level")
                {
                    return GeometricReference;
                }

                // For mullions, prefer centerline
                if (ItemType == "Mullion" || IsMullion)
                {
                    return CenterlineReference ?? GeometricReference;
                }

                // Handle explicit reference types for walls and elements
                switch (referenceType)
                {
                    case DimensionReferenceType.Centerline:
                        return GetCenterlineReference();

                    case DimensionReferenceType.ExteriorFace:
                        return GetExteriorFaceReference();

                    case DimensionReferenceType.InteriorFace:
                        return GetInteriorFaceReference();

                    case DimensionReferenceType.Auto:
                        return GetAutoReference();

                    default:
                        return GeometricReference ?? CenterlineReference;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting reference for {ItemType} {Element.Id}: {ex.Message}");
                return GeometricReference ?? CenterlineReference;
            }
        }

        /// <summary>
        /// Gets centerline reference
        /// </summary>
        private Reference GetCenterlineReference()
        {
            return CenterlineReference ?? GeometricReference;
        }

        private Reference GetExteriorFaceReference()
        {
            // If we have a pre-calculated exterior face reference, use it
            if (ExteriorFaceReference != null)
                return ExteriorFaceReference;

            // Try to get exterior face reference for walls
            if (Element is Wall wall)
            {
                var faceRef = GetWallFaceReference(wall, true); // true = exterior
                if (faceRef != null)
                    return faceRef;
            }

            // Fallback to centerline
            return GetCenterlineReference();
        }

        /// <summary>
        /// Gets interior face reference with enhanced face detection
        /// </summary>
        private Reference GetInteriorFaceReference()
        {
            // If we have a pre-calculated interior face reference, use it
            if (InteriorFaceReference != null)
                return InteriorFaceReference;

            // Try to get interior face reference for walls
            if (Element is Wall wall)
            {
                var faceRef = GetWallFaceReference(wall, false); // false = interior
                if (faceRef != null)
                    return faceRef;
            }

            // Fallback to centerline
            return GetCenterlineReference();
        }

        private Reference GetWallFaceReference(Wall wall, bool isExterior)
        {
            try
            {
                Options geometryOptions = new Options();
                geometryOptions.ComputeReferences = true;
                geometryOptions.IncludeNonVisibleObjects = false;

                GeometryElement geometryElement = wall.get_Geometry(geometryOptions);
                if (geometryElement == null) return null;

                // Get wall location curve to determine orientation
                if (!(wall.Location is LocationCurve locationCurve)) return null;
                var wallCurve = locationCurve.Curve;
                var wallDirection = (wallCurve.GetEndPoint(1) - wallCurve.GetEndPoint(0)).Normalize();
                var wallNormal = XYZ.BasisZ.CrossProduct(wallDirection).Normalize();

                // If isExterior is false, flip the normal for interior face
                if (!isExterior)
                    wallNormal = wallNormal.Negate();

                foreach (GeometryObject geoObj in geometryElement)
                {
                    if (geoObj is Solid solid)
                    {
                        foreach (Face face in solid.Faces)
                        {
                            if (face is PlanarFace planarFace && face.Reference != null)
                            {
                                var faceNormal = planarFace.FaceNormal;

                                // Check if this face is oriented correctly for exterior/interior
                                double dotProduct = faceNormal.DotProduct(wallNormal);
                                if (Math.Abs(dotProduct) > 0.7) // Face is aligned with wall normal
                                {
                                    return face.Reference;
                                }
                            }
                        }
                    }
                }

                return null;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting wall face reference: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// Automatically selects the best reference based on element type
        /// </summary>
        private Reference GetAutoReference()
        {
            // For curtain wall grid lines, always use centerline (most precise)
            if (ItemType == "CurtainGridLine")
                return CenterlineReference ?? GeometricReference;

            // For curtain walls and mullions, prefer centerline
            if (IsCurtainWallElement || IsMullion)
                return CenterlineReference ?? GeometricReference;

            // For structural walls, prefer exterior face
            if (IsStructuralWall())
                return GetExteriorFaceReference();

            // For grids and levels, always use centerline/main reference
            if (ItemType == "Grid" || ItemType == "Level")
                return GeometricReference;

            // For regular building elements, prefer centerline
            return GetCenterlineReference();
        }

        /// <summary>
        /// Checks if this element is a structural wall
        /// </summary>
        private bool IsStructuralWall()
        {
            if (Element?.Category?.Name == null) return false;

            string categoryName = Element.Category.Name.ToLower();

            // Check category
            if (categoryName.Contains("structural"))
                return true;

            // Check if it's a wall with structural function
            if (Element is Wall wall)
            {
                try
                {
                    var wallType = wall.WallType;
                    if (wallType != null)
                    {
                        string typeName = wallType.Name.ToLower();
                        return typeName.Contains("structural") ||
                               typeName.Contains("bearing") ||
                               typeName.Contains("shear");
                    }
                }
                catch
                {
                    // Ignore errors in wall type access
                }
            }

            return false;
        }

        /// <summary>
        /// Gets display name for debugging/logging
        /// </summary>
        public string GetDisplayName()
        {
            try
            {
                string elementName = Element?.Name ?? "Unnamed";
                string elementId = Element?.Id?.IntegerValue.ToString() ?? "No ID";

                if (ItemType == "CurtainGridLine")
                {
                    return $"Curtain Grid Line {elementId} ({elementName}) - Parent: {ParentWallId?.IntegerValue}";
                }
                else if (ItemType == "Mullion")
                {
                    return $"Mullion {elementId} ({elementName}) - Parent: {ParentWallId?.IntegerValue}";
                }
                else if (ItemType == "CurtainWall")
                {
                    return $"Curtain Wall {elementId} ({elementName})";
                }
                else if (ItemType == "Grid")
                {
                    return $"Grid {elementId} ({elementName})";
                }
                else if (ItemType == "Level")
                {
                    return $"Level {elementId} ({elementName})";
                }
                else
                {
                    return $"{ItemType} {elementId} ({elementName})";
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting display name: {ex.Message}");
                return $"{ItemType ?? "Unknown"} - {Element?.Id?.IntegerValue ?? 0}";
            }
        }

        /// <summary>
        /// Validates that this projected item has valid references for dimensioning
        /// </summary>
        public bool IsValidForDimensioning()
        {
            try
            {
                // Must have an element
                if (Element == null) return false;

                // Must have at least one valid reference
                if (GeometricReference == null && CenterlineReference == null) return false;

                // For linear elements, must have valid direction or be marked as point element
                if (!IsPointElement && (ProjectedDirection == null || ProjectedDirection.GetLength() < 0.001))
                    return false;

                // For point elements, must have valid point
                if (IsPointElement && ProjectedPoint == null) return false;

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error validating projected item: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Gets the projected point adjusted for the specified reference type
        /// </summary>
        public XYZ GetAdjustedProjectedPoint(DimensionReferenceType referenceType, XYZ viewDirection = null)
        {
            if (referenceType == DimensionReferenceType.Centerline ||
                referenceType == DimensionReferenceType.Auto ||
                ElementWidth <= 0.001)
            {
                return ProjectedPoint;
            }

            // Calculate face offset for walls
            if (viewDirection != null && ElementWidth > 0.001)
            {
                XYZ perpendicular = GetPerpendicularDirection(viewDirection);
                double offset = ElementWidth / 2.0;

                if (referenceType == DimensionReferenceType.ExteriorFace)
                {
                    return ProjectedPoint + perpendicular * offset;
                }
                else if (referenceType == DimensionReferenceType.InteriorFace)
                {
                    return ProjectedPoint - perpendicular * offset;
                }
            }

            return ProjectedPoint;
        }

        /// <summary>
        /// Gets perpendicular direction for face offset calculation
        /// </summary>
        private XYZ GetPerpendicularDirection(XYZ direction)
        {
            if (direction.IsAlmostEqualTo(XYZ.BasisZ))
            {
                return XYZ.BasisX;
            }

            XYZ perpendicular = direction.CrossProduct(XYZ.BasisZ);
            if (perpendicular.GetLength() < 0.001)
            {
                perpendicular = direction.CrossProduct(XYZ.BasisY);
            }

            return perpendicular.Normalize();
        }
    }
}