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
                switch (referenceType)
                {
                    case DimensionReferenceType.Centerline:
                        return CenterlineReference ?? GeometricReference;

                    case DimensionReferenceType.ExteriorFace:
                        return ExteriorFaceReference ?? CenterlineReference ?? GeometricReference;

                    case DimensionReferenceType.InteriorFace:
                        return InteriorFaceReference ?? CenterlineReference ?? GeometricReference;

                    default:
                        return GeometricReference;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting reference: {ex.Message}");
                return GeometricReference;
            }
        }

        /// <summary>
        /// Automatically selects the best reference based on element type
        /// </summary>
        private Reference GetAutoReference()
        {
            // For curtain walls and mullions, prefer centerline
            if (IsCurtainWallElement || IsMullion)
            {
                return CenterlineReference ?? GeometricReference;
            }

            // For structural walls, prefer exterior face
            if (IsStructuralWall())
            {
                return ExteriorFaceReference ?? CenterlineReference ?? GeometricReference;
            }

            // For grids and levels, always use centerline/main reference
            if (ItemType == "Grid" || ItemType == "Level")
            {
                return GeometricReference;
            }

            // For other elements, prefer centerline
            return CenterlineReference ?? GeometricReference;
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