using System;
using System.IO;
using System.Xml.Serialization;
using Autodesk.Revit.DB;

namespace CustomRevitCommand
{
    [Serializable]
    public class DimensionSettings
    {
        public double DefaultOffset { get; set; } = 1.64; // feet
        public bool IncludeGrids { get; set; } = true;
        public bool IncludeLevels { get; set; } = true;
        public bool IncludeStructural { get; set; } = true;
        public bool IncludeCurtainWalls { get; set; } = true;
        public bool IncludeMullions { get; set; } = true;
        public double CollinearityTolerance { get; set; } = 0.01;
        public double PerpendicularTolerance { get; set; } = 0.1;
        public DimensionReferenceType ReferenceType { get; set; } = DimensionReferenceType.Centerline;
        public bool ShowProgressDialog { get; set; } = true;
        public bool AutoSelectCurrentView { get; set; } = true;
        public bool CreateOverallDimensions { get; set; } = false;

        // Adaptive tolerance settings
        public double StructuralTolerance { get; set; } = 0.05; // 15mm
        public double GridTolerance { get; set; } = 0.005; // 1.5mm
        public double CurtainWallTolerance { get; set; } = 0.008; // 2.5mm

        public static DimensionSettings LoadFromProject(Document doc)
        {
            try
            {
                // Try to load from project parameters first
                var settings = LoadFromProjectParameters(doc);
                if (settings != null) return settings;

                // Fallback to user settings file
                return LoadFromUserSettings() ?? new DimensionSettings();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading settings: {ex.Message}");
                return new DimensionSettings();
            }
        }

        private static DimensionSettings LoadFromProjectParameters(Document doc)
        {
            try
            {
                // Look for project parameter storing settings
                var projectInfo = doc.ProjectInformation;

                // Try to get parameter by name (this is a simplified approach)
                // In a full implementation, you'd create and manage shared parameters
                foreach (Parameter param in projectInfo.Parameters)
                {
                    if (param.Definition.Name == "DEAXO_DimensionSettings" && param.HasValue)
                    {
                        string settingsXml = param.AsString();
                        return DeserializeFromXml(settingsXml);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading from project parameters: {ex.Message}");
            }
            return null;
        }

        private static DimensionSettings LoadFromUserSettings()
        {
            try
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string settingsPath = Path.Combine(appDataPath, "DEAXO", "DimensionSettings.xml");

                if (File.Exists(settingsPath))
                {
                    string xml = File.ReadAllText(settingsPath);
                    return DeserializeFromXml(xml);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading user settings: {ex.Message}");
            }
            return null;
        }

        public void SaveToProject(Document doc)
        {
            try
            {
                // Save to project parameters
                using (var trans = new Transaction(doc, "Save Dimension Settings"))
                {
                    trans.Start();

                    var projectInfo = doc.ProjectInformation;

                    // Try to find existing parameter
                    Parameter param = null;
                    foreach (Parameter p in projectInfo.Parameters)
                    {
                        if (p.Definition.Name == "DEAXO_DimensionSettings")
                        {
                            param = p;
                            break;
                        }
                    }

                    if (param == null)
                    {
                        // Create parameter if it doesn't exist (simplified approach)
                        System.Diagnostics.Debug.WriteLine("Settings parameter not found - using user settings only");
                        trans.RollBack();
                        SaveToUserSettings();
                        return;
                    }

                    if (param != null)
                    {
                        string xml = SerializeToXml();
                        param.Set(xml);
                    }

                    trans.Commit();
                }

                // Also save to user settings as backup
                SaveToUserSettings();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving settings to project: {ex.Message}");
                // Fallback to user settings only
                SaveToUserSettings();
            }
        }

        private void SaveToUserSettings()
        {
            try
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string deaxoPath = Path.Combine(appDataPath, "DEAXO");

                if (!Directory.Exists(deaxoPath))
                {
                    Directory.CreateDirectory(deaxoPath);
                }

                string settingsPath = Path.Combine(deaxoPath, "DimensionSettings.xml");
                string xml = SerializeToXml();
                File.WriteAllText(settingsPath, xml);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving user settings: {ex.Message}");
            }
        }

        private void CreateProjectParameter(Document doc)
        {
            try
            {
                // This would require creating a shared parameter file
                // For now, we'll skip this and rely on user settings
                System.Diagnostics.Debug.WriteLine("Project parameter creation not implemented - using user settings only");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error creating project parameter: {ex.Message}");
            }
        }

        private string SerializeToXml()
        {
            try
            {
                var serializer = new XmlSerializer(typeof(DimensionSettings));
                using (var writer = new StringWriter())
                {
                    serializer.Serialize(writer, this);
                    return writer.ToString();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error serializing settings: {ex.Message}");
                return "";
            }
        }

        private static DimensionSettings DeserializeFromXml(string xml)
        {
            try
            {
                var serializer = new XmlSerializer(typeof(DimensionSettings));
                using (var reader = new StringReader(xml))
                {
                    return (DimensionSettings)serializer.Deserialize(reader);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error deserializing settings: {ex.Message}");
                return new DimensionSettings();
            }
        }

        public double GetToleranceForElement(Element element)
        {
            if (element == null) return CollinearityTolerance;

            if (IsStructuralElement(element))
                return StructuralTolerance;

            if (element is Grid)
                return GridTolerance;

            if (IsCurtainWallElement(element))
                return CurtainWallTolerance;

            return CollinearityTolerance;
        }

        private bool IsStructuralElement(Element element)
        {
            if (element?.Category?.Name == null) return false;
            string catName = element.Category.Name.ToLower();
            return catName.Contains("structural") || catName.Contains("beam") ||
                   catName.Contains("column") || catName.Contains("wall");
        }

        private bool IsCurtainWallElement(Element element)
        {
            if (element?.Category?.Name == null) return false;
            string catName = element.Category.Name.ToLower();
            return catName.Contains("curtain") || catName.Contains("mullion") || catName.Contains("panel");
        }

        public static string GetReferenceTypeDescription(DimensionReferenceType referenceType)
        {
            switch (referenceType)
            {
                case DimensionReferenceType.Centerline:
                    return "Centerline - Dimensions to element centerlines (default for grids, levels)";
                case DimensionReferenceType.ExteriorFace:
                    return "Exterior Face - Dimensions to outer face of walls and elements";
                case DimensionReferenceType.InteriorFace:
                    return "Interior Face - Dimensions to inner face of walls and elements";
                case DimensionReferenceType.Auto:
                    return "Auto - System chooses best reference based on element type";
                default:
                    return "Unknown reference type";
            }
        }
    }

    public enum DimensionReferenceType
    {
        Centerline,
        ExteriorFace,
        InteriorFace,
        Auto // Let system decide based on element type
    }
}