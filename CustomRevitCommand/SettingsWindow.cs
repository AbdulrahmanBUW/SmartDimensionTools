using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using Autodesk.Revit.DB;

// Aliases to resolve namespace conflicts
using WpfGrid = System.Windows.Controls.Grid;
using WpfTextBox = System.Windows.Controls.TextBox;

namespace CustomRevitCommand
{
    public class SettingsWindow : Window
    {
        private DimensionSettings _settings;
        private Document _document;

        // UI Controls
        private CheckBox _includeGridsCheckBox;
        private CheckBox _includeLevelsCheckBox;
        private CheckBox _includeStructuralCheckBox;
        private CheckBox _includeCurtainWallsCheckBox;
        private CheckBox _includeMullionsCheckBox;
        private CheckBox _showProgressCheckBox;
        private CheckBox _autoSelectCurrentViewCheckBox;
        private CheckBox _createOverallDimensionsCheckBox;

        private WpfTextBox _defaultOffsetTextBox;
        private WpfTextBox _collinearityToleranceTextBox;
        private WpfTextBox _perpendicularToleranceTextBox;
        private WpfTextBox _structuralToleranceTextBox;
        private WpfTextBox _gridToleranceTextBox;
        private WpfTextBox _curtainWallToleranceTextBox;

        private ComboBox _referenceTypeComboBox;

        private Button _okButton;
        private Button _cancelButton;
        private Button _resetButton;

        public DimensionSettings Settings => _settings;

        public SettingsWindow(Document document, DimensionSettings currentSettings = null)
        {
            _document = document;
            _settings = currentSettings?.Clone() ?? DimensionSettings.LoadFromProject(document);

            InitializeComponent();
            LoadSettingsToUI();
        }

        private void InitializeComponent()
        {
            Title = "Auto-Dimension Settings";
            Height = 600;
            Width = 450;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.CanResize;
            MinHeight = 500;
            MinWidth = 400;

            var mainGrid = new WpfGrid();
            mainGrid.Margin = new Thickness(10);

            // Define rows
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Title
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) }); // Content
            mainGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto }); // Buttons

            // Title
            var titleText = new TextBlock
            {
                Text = "Auto-Dimension Settings",
                FontSize = 16,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 15),
                HorizontalAlignment = HorizontalAlignment.Center
            };
            WpfGrid.SetRow(titleText, 0);
            mainGrid.Children.Add(titleText);

            // Content in ScrollViewer
            var scrollViewer = new ScrollViewer
            {
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Disabled
            };

            var contentPanel = CreateContentPanel();
            scrollViewer.Content = contentPanel;
            WpfGrid.SetRow(scrollViewer, 1);
            mainGrid.Children.Add(scrollViewer);

            // Buttons
            var buttonPanel = CreateButtonPanel();
            WpfGrid.SetRow(buttonPanel, 2);
            mainGrid.Children.Add(buttonPanel);

            Content = mainGrid;
        }

        private StackPanel CreateContentPanel()
        {
            var panel = new StackPanel();

            // Element Inclusion Section
            panel.Children.Add(CreateSectionHeader("Element Inclusion"));

            _includeGridsCheckBox = CreateCheckBox("Include Grids", "Automatically include grid lines in dimensioning");
            panel.Children.Add(_includeGridsCheckBox);

            _includeLevelsCheckBox = CreateCheckBox("Include Levels", "Automatically include levels in section/elevation views");
            panel.Children.Add(_includeLevelsCheckBox);

            _includeStructuralCheckBox = CreateCheckBox("Include Structural Elements", "Include beams, columns, structural walls");
            panel.Children.Add(_includeStructuralCheckBox);

            _includeCurtainWallsCheckBox = CreateCheckBox("Include Curtain Walls", "Include curtain wall panels and systems");
            panel.Children.Add(_includeCurtainWallsCheckBox);

            _includeMullionsCheckBox = CreateCheckBox("Include Curtain Wall Grid Lines", "Include curtain wall grid lines for precise dimensioning");
            panel.Children.Add(_includeMullionsCheckBox);

            // Reference Type Section
            panel.Children.Add(CreateSectionHeader("Dimension Reference"));

            var referencePanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 5, 0, 10) };
            referencePanel.Children.Add(new TextBlock { Text = "Reference Type:", VerticalAlignment = VerticalAlignment.Center, Width = 120 });

            _referenceTypeComboBox = new ComboBox { Width = 150, Margin = new Thickness(10, 0, 0, 0) };
            _referenceTypeComboBox.Items.Add(new ComboBoxItem { Content = "Centerline", Tag = DimensionReferenceType.Centerline });
            _referenceTypeComboBox.Items.Add(new ComboBoxItem { Content = "Exterior Face", Tag = DimensionReferenceType.ExteriorFace });
            _referenceTypeComboBox.Items.Add(new ComboBoxItem { Content = "Interior Face", Tag = DimensionReferenceType.InteriorFace });
            _referenceTypeComboBox.Items.Add(new ComboBoxItem { Content = "Auto (Recommended)", Tag = DimensionReferenceType.Auto });
            referencePanel.Children.Add(_referenceTypeComboBox);
            panel.Children.Add(referencePanel);

            // Add explanation for reference types
            var explanationText = new TextBlock
            {
                Text = "• Centerline: Dimension to element centerlines\n" +
                       "• Exterior Face: Dimension to outer face of walls/elements\n" +
                       "• Interior Face: Dimension to inner face of walls/elements\n" +
                       "• Auto: System chooses best reference based on element type\n\n" +
                       "Note: Curtain wall grid lines always use centerline reference",
                FontSize = 10,
                Foreground = Brushes.Gray,
                Margin = new Thickness(0, 0, 0, 15),
                TextWrapping = TextWrapping.Wrap
            };
            panel.Children.Add(explanationText);

            // Tolerances Section
            panel.Children.Add(CreateSectionHeader("Tolerances & Spacing"));

            _defaultOffsetTextBox = CreateLabeledTextBox(panel, "Default Offset (feet):", "Distance from elements to dimension line");
            _collinearityToleranceTextBox = CreateLabeledTextBox(panel, "Collinearity Tolerance (feet):", "How close elements must be to be considered on same line");
            _perpendicularToleranceTextBox = CreateLabeledTextBox(panel, "Perpendicular Tolerance:", "Tolerance for detecting perpendicular elements");

            // Advanced Tolerances
            panel.Children.Add(CreateSectionHeader("Advanced Tolerances"));
            _structuralToleranceTextBox = CreateLabeledTextBox(panel, "Structural Elements (feet):", "Relaxed tolerance for structural elements");
            _gridToleranceTextBox = CreateLabeledTextBox(panel, "Grid Lines (feet):", "Tight tolerance for grid alignment");
            _curtainWallToleranceTextBox = CreateLabeledTextBox(panel, "Curtain Wall Grid (feet):", "Tolerance for curtain wall grid line alignment");

            // Behavior Section
            panel.Children.Add(CreateSectionHeader("Behavior"));

            _showProgressCheckBox = CreateCheckBox("Show Progress Dialog", "Display progress when processing multiple views");
            panel.Children.Add(_showProgressCheckBox);

            _autoSelectCurrentViewCheckBox = CreateCheckBox("Auto-select Current View", "Automatically select the active view in view selection dialog");
            panel.Children.Add(_autoSelectCurrentViewCheckBox);

            _createOverallDimensionsCheckBox = CreateCheckBox("Create Overall Dimensions", "Create overall building dimensions in addition to individual chains");
            panel.Children.Add(_createOverallDimensionsCheckBox);

            return panel;
        }

        private TextBlock CreateSectionHeader(string text)
        {
            return new TextBlock
            {
                Text = text,
                FontWeight = FontWeights.Bold,
                FontSize = 12,
                Margin = new Thickness(0, 15, 0, 5),
                Foreground = Brushes.DarkBlue
            };
        }

        private CheckBox CreateCheckBox(string content, string tooltip)
        {
            return new CheckBox
            {
                Content = content,
                Margin = new Thickness(0, 2, 0, 2),
                ToolTip = tooltip
            };
        }

        private WpfTextBox CreateLabeledTextBox(StackPanel parent, string label, string tooltip)
        {
            var panel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 2, 0, 2) };

            var labelBlock = new TextBlock
            {
                Text = label,
                Width = 180,
                VerticalAlignment = VerticalAlignment.Center,
                ToolTip = tooltip
            };
            panel.Children.Add(labelBlock);

            var textBox = new WpfTextBox
            {
                Width = 100,
                Height = 23,
                Margin = new Thickness(5, 0, 0, 0),
                ToolTip = tooltip
            };
            panel.Children.Add(textBox);

            parent.Children.Add(panel);
            return textBox;
        }

        private StackPanel CreateButtonPanel()
        {
            var panel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 15, 0, 0)
            };

            _resetButton = new Button
            {
                Content = "Reset to Defaults",
                Padding = new Thickness(15, 5, 15, 5),
                Margin = new Thickness(0, 0, 10, 0),
                MinWidth = 100
            };
            _resetButton.Click += ResetButton_Click;
            panel.Children.Add(_resetButton);

            _okButton = new Button
            {
                Content = "OK",
                Padding = new Thickness(15, 5, 15, 5),
                Margin = new Thickness(5, 0, 5, 0),
                MinWidth = 80,
                IsDefault = true
            };
            _okButton.Click += OkButton_Click;
            panel.Children.Add(_okButton);

            _cancelButton = new Button
            {
                Content = "Cancel",
                Padding = new Thickness(15, 5, 15, 5),
                Margin = new Thickness(5, 0, 0, 0),
                MinWidth = 80,
                IsCancel = true
            };
            _cancelButton.Click += CancelButton_Click;
            panel.Children.Add(_cancelButton);

            return panel;
        }

        private void LoadSettingsToUI()
        {
            try
            {
                _includeGridsCheckBox.IsChecked = _settings.IncludeGrids;
                _includeLevelsCheckBox.IsChecked = _settings.IncludeLevels;
                _includeStructuralCheckBox.IsChecked = _settings.IncludeStructural;
                _includeCurtainWallsCheckBox.IsChecked = _settings.IncludeCurtainWalls;
                _includeMullionsCheckBox.IsChecked = _settings.IncludeMullions;
                _showProgressCheckBox.IsChecked = _settings.ShowProgressDialog;
                _autoSelectCurrentViewCheckBox.IsChecked = _settings.AutoSelectCurrentView;
                _createOverallDimensionsCheckBox.IsChecked = _settings.CreateOverallDimensions;

                _defaultOffsetTextBox.Text = _settings.DefaultOffset.ToString("F3");
                _collinearityToleranceTextBox.Text = _settings.CollinearityTolerance.ToString("F4");
                _perpendicularToleranceTextBox.Text = _settings.PerpendicularTolerance.ToString("F3");
                _structuralToleranceTextBox.Text = _settings.StructuralTolerance.ToString("F4");
                _gridToleranceTextBox.Text = _settings.GridTolerance.ToString("F4");
                _curtainWallToleranceTextBox.Text = _settings.CurtainWallTolerance.ToString("F4");

                // Set reference type combo box
                foreach (ComboBoxItem item in _referenceTypeComboBox.Items)
                {
                    if ((DimensionReferenceType)item.Tag == _settings.ReferenceType)
                    {
                        _referenceTypeComboBox.SelectedItem = item;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading settings to UI: {ex.Message}");
            }
        }

        private bool SaveUIToSettings()
        {
            try
            {
                _settings.IncludeGrids = _includeGridsCheckBox.IsChecked ?? true;
                _settings.IncludeLevels = _includeLevelsCheckBox.IsChecked ?? true;
                _settings.IncludeStructural = _includeStructuralCheckBox.IsChecked ?? true;
                _settings.IncludeCurtainWalls = _includeCurtainWallsCheckBox.IsChecked ?? true;
                _settings.IncludeMullions = _includeMullionsCheckBox.IsChecked ?? true;
                _settings.ShowProgressDialog = _showProgressCheckBox.IsChecked ?? true;
                _settings.AutoSelectCurrentView = _autoSelectCurrentViewCheckBox.IsChecked ?? true;
                _settings.CreateOverallDimensions = _createOverallDimensionsCheckBox.IsChecked ?? false;

                if (!double.TryParse(_defaultOffsetTextBox.Text, out double defaultOffset) || defaultOffset <= 0)
                {
                    MessageBox.Show("Invalid default offset value. Please enter a positive number.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
                _settings.DefaultOffset = defaultOffset;

                if (!double.TryParse(_collinearityToleranceTextBox.Text, out double collinearTol) || collinearTol <= 0)
                {
                    MessageBox.Show("Invalid collinearity tolerance value. Please enter a positive number.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
                _settings.CollinearityTolerance = collinearTol;

                if (!double.TryParse(_perpendicularToleranceTextBox.Text, out double perpTol) || perpTol <= 0)
                {
                    MessageBox.Show("Invalid perpendicular tolerance value. Please enter a positive number.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
                _settings.PerpendicularTolerance = perpTol;

                if (!double.TryParse(_structuralToleranceTextBox.Text, out double structTol) || structTol <= 0)
                {
                    MessageBox.Show("Invalid structural tolerance value. Please enter a positive number.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
                _settings.StructuralTolerance = structTol;

                if (!double.TryParse(_gridToleranceTextBox.Text, out double gridTol) || gridTol <= 0)
                {
                    MessageBox.Show("Invalid grid tolerance value. Please enter a positive number.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
                _settings.GridTolerance = gridTol;

                if (!double.TryParse(_curtainWallToleranceTextBox.Text, out double cwTol) || cwTol <= 0)
                {
                    MessageBox.Show("Invalid curtain wall tolerance value. Please enter a positive number.", "Invalid Input", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return false;
                }
                _settings.CurtainWallTolerance = cwTol;

                if (_referenceTypeComboBox.SelectedItem is ComboBoxItem selectedItem)
                {
                    _settings.ReferenceType = (DimensionReferenceType)selectedItem.Tag;
                }

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving settings: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return false;
            }
        }

        private void ResetButton_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("Reset all settings to default values?", "Reset Settings",
                MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                _settings = new DimensionSettings();
                LoadSettingsToUI();
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            if (SaveUIToSettings())
            {
                try
                {
                    _settings.SaveToProject(_document);
                    DialogResult = true;
                    Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Settings saved locally but could not save to project: {ex.Message}",
                        "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    DialogResult = true;
                    Close();
                }
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }

    // Extension method to clone settings
    public static class DimensionSettingsExtensions
    {
        public static DimensionSettings Clone(this DimensionSettings original)
        {
            return new DimensionSettings
            {
                DefaultOffset = original.DefaultOffset,
                IncludeGrids = original.IncludeGrids,
                IncludeLevels = original.IncludeLevels,
                IncludeStructural = original.IncludeStructural,
                IncludeCurtainWalls = original.IncludeCurtainWalls,
                IncludeMullions = original.IncludeMullions,
                CollinearityTolerance = original.CollinearityTolerance,
                PerpendicularTolerance = original.PerpendicularTolerance,
                ReferenceType = original.ReferenceType,
                ShowProgressDialog = original.ShowProgressDialog,
                AutoSelectCurrentView = original.AutoSelectCurrentView,
                CreateOverallDimensions = original.CreateOverallDimensions,
                StructuralTolerance = original.StructuralTolerance,
                GridTolerance = original.GridTolerance,
                CurtainWallTolerance = original.CurtainWallTolerance
            };
        }
    }
}