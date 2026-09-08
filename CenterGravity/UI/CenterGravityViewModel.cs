using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Windows;
using System.Windows.Input;

namespace BBI.JD.UI
{
    public class ReferenceOption
    {
        public ReferenceMode Mode { get; set; }
        public string Label { get; set; }
        public override string ToString() => Label;
    }

    public class WeightOption
    {
        public WeightMode Mode { get; set; }
        public string Label { get; set; }
        public override string ToString() => Label;
    }

    public class CenterGravityViewModel : INotifyPropertyChanged
    {
        private Units units;
        private ForgeTypeId lengthUnit;
        private ForgeTypeId volumeUnit;
        private ForgeTypeId massUnit;

        private XYZ centroidInternal = XYZ.Zero;
        private XYZ massCentroidInternal = XYZ.Zero;
        private XYZ projectBasePoint = XYZ.Zero;
        private XYZ surveyPoint = XYZ.Zero;
        private double volumeInternal;
        private double massInternal;
        private bool hasResult;
        private bool massValid;
        private bool massComplete;
        private int noDensityCount;
        private bool hasMarkers;
        private int skippedCount;
        private int expandedContainers;

        public CenterGravityViewModel()
        {
            ReferenceOptions = new[]
            {
                new ReferenceOption { Mode = ReferenceMode.InternalOrigin,   Label = "Internal origin" },
                new ReferenceOption { Mode = ReferenceMode.ProjectBasePoint, Label = "Project base point" },
                new ReferenceOption { Mode = ReferenceMode.SurveyPoint,      Label = "Survey point" },
            };
            selectedReference = ReferenceOptions[0];

            WeightOptions = new[]
            {
                new WeightOption { Mode = WeightMode.Volume, Label = "Volume" },
                new WeightOption { Mode = WeightMode.Mass,   Label = "Mass (material density)" },
            };
            selectedWeight = WeightOptions[0];

            PlaceCommand = new RelayCommand(() => PlaceRequested?.Invoke(), () => hasResult);
            ClearCommand = new RelayCommand(() => ClearRequested?.Invoke(), () => hasMarkers);
            CopyCommand = new RelayCommand(CopyCoordinates, () => hasResult);
            ExportCsvCommand = new RelayCommand(ExportCsv, () => hasResult);
        }

        /// <summary>Raised when the user asks to drop the centre-of-gravity marker into the model.</summary>
        public Action PlaceRequested;

        /// <summary>Raised when the user asks to remove the drawn centre-of-gravity marker(s).</summary>
        public Action ClearRequested;

        /// <summary>Raised when the default density changes; the pane pushes it to the handler and recomputes.</summary>
        public Action DensityChanged;

        /// <summary>Raised when the volume/mass weighting changes (no recompute needed - both are cached).</summary>
        public Action WeightModeChanged;

        public ObservableCollection<CgRow> Rows { get; } = new ObservableCollection<CgRow>();

        public ReferenceOption[] ReferenceOptions { get; }
        public WeightOption[] WeightOptions { get; }

        private ReferenceOption selectedReference;
        public ReferenceOption SelectedReference
        {
            get => selectedReference;
            set
            {
                if (value == null || ReferenceEquals(value, selectedReference))
                {
                    return;
                }

                selectedReference = value;
                OnPropertyChanged();
                RefreshCoordinates();
            }
        }

        private WeightOption selectedWeight;
        public WeightOption SelectedWeight
        {
            get => selectedWeight;
            set
            {
                if (value == null || ReferenceEquals(value, selectedWeight))
                {
                    return;
                }

                selectedWeight = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(WeightByMass));
                RefreshCoordinates();
                WeightModeChanged?.Invoke();
            }
        }

        public bool WeightByMass => selectedWeight?.Mode == WeightMode.Mass;

        private string defaultDensityText = string.Empty;
        public string DefaultDensityText
        {
            get => defaultDensityText;
            set
            {
                if (value == defaultDensityText)
                {
                    return;
                }

                defaultDensityText = value;
                OnPropertyChanged();
                ParseDensity();
                DensityChanged?.Invoke();
            }
        }

        /// <summary>Fallback density in internal units, or null when the box is empty / unparseable.</summary>
        public double? DefaultDensityInternal { get; private set; }

        public string DensityUnitLabel { get; private set; } = string.Empty;

        public ICommand PlaceCommand { get; }
        public ICommand ClearCommand { get; }
        public ICommand CopyCommand { get; }
        public ICommand ExportCsvCommand { get; }

        public bool HasResult
        {
            get => hasResult;
            private set { hasResult = value; OnPropertyChanged(); }
        }

        public bool HasMarkers
        {
            get => hasMarkers;
            set { hasMarkers = value; OnPropertyChanged(); CommandManager.InvalidateRequerySuggested(); }
        }

        public string Volume { get; private set; } = "-";
        public string Mass { get; private set; } = "-";
        public string X { get; private set; } = "-";
        public string Y { get; private set; } = "-";
        public string Z { get; private set; } = "-";
        public string Xyz { get; private set; } = "-";

        public string Warning { get; private set; } = string.Empty;
        public bool HasWarning => !string.IsNullOrEmpty(Warning);

        public string Summary
        {
            get
            {
                if (Rows.Count == 0)
                {
                    return "Select model elements in Revit.";
                }

                int used = Rows.Count - skippedCount;
                string text = string.Format("{0} element(s), {1} in the calculation", Rows.Count, used);

                if (skippedCount > 0)
                {
                    text += string.Format(" - {0} without solid geometry", skippedCount);
                }

                if (expandedContainers > 0)
                {
                    text += string.Format(" - expanded from {0} group/assembly", expandedContainers);
                }

                return text;
            }
        }

        /// <summary>Feed a fresh result in (called on the UI thread by the pane).</summary>
        public void SetResult(Units units, CentroidVolume cv, IEnumerable<CgRow> rows,
            XYZ projectBasePoint, XYZ surveyPoint, int expandedContainers)
        {
            this.units = units;
            lengthUnit = units?.GetFormatOptions(SpecTypeId.Length)?.GetUnitTypeId();
            volumeUnit = units?.GetFormatOptions(SpecTypeId.Volume)?.GetUnitTypeId();
            massUnit = units?.GetFormatOptions(SpecTypeId.Mass)?.GetUnitTypeId();

            UpdateDensityUnitLabel();

            this.projectBasePoint = projectBasePoint ?? XYZ.Zero;
            this.surveyPoint = surveyPoint ?? XYZ.Zero;
            this.expandedContainers = expandedContainers;

            Rows.Clear();
            if (rows != null)
            {
                foreach (CgRow r in rows)
                {
                    Rows.Add(r);
                }
            }

            skippedCount = Rows.Count(r => r.Skipped);

            hasResult = cv != null && cv.IsValid;
            centroidInternal = hasResult ? cv.Centroid : XYZ.Zero;
            volumeInternal = hasResult ? cv.Volume : 0.0;

            massValid = cv != null && cv.MassIsValid;
            massComplete = cv != null && cv.MassComplete;
            massCentroidInternal = massValid ? cv.MassCentroid : XYZ.Zero;
            massInternal = massValid ? cv.Mass : 0.0;
            noDensityCount = cv?.NoDensityElementIds.Count ?? 0;

            OnPropertyChanged(nameof(HasResult));
            OnPropertyChanged(nameof(Summary));
            RefreshCoordinates();

            CommandManager.InvalidateRequerySuggested();
        }

        public void Reset()
        {
            Rows.Clear();
            skippedCount = 0;
            expandedContainers = 0;
            hasResult = false;
            centroidInternal = XYZ.Zero;
            massCentroidInternal = XYZ.Zero;
            volumeInternal = 0.0;
            massInternal = 0.0;
            massValid = false;
            massComplete = false;
            noDensityCount = 0;

            OnPropertyChanged(nameof(HasResult));
            OnPropertyChanged(nameof(Summary));
            RefreshCoordinates();

            CommandManager.InvalidateRequerySuggested();
        }

        private XYZ CurrentOrigin()
        {
            switch (selectedReference?.Mode)
            {
                case ReferenceMode.ProjectBasePoint: return projectBasePoint;
                case ReferenceMode.SurveyPoint: return surveyPoint;
                default: return XYZ.Zero;
            }
        }

        private void RefreshCoordinates()
        {
            Warning = string.Empty;

            bool massMode = WeightByMass;
            bool showMass = massMode && massValid;

            if (!hasResult)
            {
                Volume = Mass = X = Y = Z = Xyz = "-";
            }
            else
            {
                Volume = FormatValue(volumeInternal, volumeUnit, SpecTypeId.Volume);
                Mass = massValid ? FormatValue(massInternal, massUnit, SpecTypeId.Mass) : "-";

                XYZ source = showMass ? massCentroidInternal : centroidInternal;

                if (massMode && !massValid)
                {
                    X = Y = Z = Xyz = "-";
                    Warning = "No material density found - enter a default density to weight by mass.";
                }
                else
                {
                    XYZ p = source - CurrentOrigin();
                    X = FormatValue(p.X, lengthUnit, SpecTypeId.Length);
                    Y = FormatValue(p.Y, lengthUnit, SpecTypeId.Length);
                    Z = FormatValue(p.Z, lengthUnit, SpecTypeId.Length);
                    Xyz = string.Format("{0}; {1}; {2}", X, Y, Z);

                    if (massMode && !massComplete && noDensityCount > 0)
                    {
                        Warning = string.Format("{0} element(s) without density are excluded from the mass result.", noDensityCount);
                    }
                }
            }

            OnPropertyChanged(nameof(Volume));
            OnPropertyChanged(nameof(Mass));
            OnPropertyChanged(nameof(X));
            OnPropertyChanged(nameof(Y));
            OnPropertyChanged(nameof(Z));
            OnPropertyChanged(nameof(Xyz));
            OnPropertyChanged(nameof(Warning));
            OnPropertyChanged(nameof(HasWarning));
        }

        private string FormatValue(double internalValue, ForgeTypeId unit, ForgeTypeId spec)
        {
            if (units != null && spec != null)
            {
                try
                {
                    return UnitFormatUtils.Format(units, spec, internalValue, false);
                }
                catch (Exception)
                {
                    // fall through to the raw conversion
                }
            }

            double display = unit != null
                ? UnitUtils.ConvertFromInternalUnits(internalValue, unit)
                : internalValue;

            return display.ToString("0.###", CultureInfo.CurrentCulture);
        }

        private void UpdateDensityUnitLabel()
        {
            try
            {
                ForgeTypeId densityUnit = units?.GetFormatOptions(SpecTypeId.MassDensity)?.GetUnitTypeId();
                DensityUnitLabel = densityUnit != null ? LabelUtils.GetLabelForUnit(densityUnit) : string.Empty;
            }
            catch (Exception)
            {
                DensityUnitLabel = string.Empty;
            }

            OnPropertyChanged(nameof(DensityUnitLabel));
        }

        private void ParseDensity()
        {
            if (string.IsNullOrWhiteSpace(defaultDensityText))
            {
                DefaultDensityInternal = null;
                return;
            }

            if (units != null &&
                UnitFormatUtils.TryParse(units, SpecTypeId.MassDensity, defaultDensityText, out double internalValue) &&
                internalValue > 0)
            {
                DefaultDensityInternal = internalValue;
            }
            else if (double.TryParse(defaultDensityText, NumberStyles.Any, CultureInfo.CurrentCulture, out double raw) && raw > 0)
            {
                ForgeTypeId densityUnit = units?.GetFormatOptions(SpecTypeId.MassDensity)?.GetUnitTypeId();
                DefaultDensityInternal = densityUnit != null
                    ? UnitUtils.ConvertToInternalUnits(raw, densityUnit)
                    : raw;
            }
            else
            {
                DefaultDensityInternal = null;
            }
        }

        private void CopyCoordinates()
        {
            try
            {
                Clipboard.SetText(Xyz ?? string.Empty);
            }
            catch (Exception)
            {
                // clipboard can be locked by another process; ignore
            }
        }

        private void ExportCsv()
        {
            Microsoft.Win32.SaveFileDialog dialog = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Export centre of gravity",
                Filter = "CSV file (*.csv)|*.csv",
                FileName = "CenterGravity.csv",
                AddExtension = true
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            string sep = CultureInfo.CurrentCulture.TextInfo.ListSeparator;
            if (sep == ".") sep = ",";

            bool massMode = WeightByMass && massValid;

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Join(sep, "Id", "Category", "Name", "Volume", "Mass"));

            foreach (CgRow r in Rows)
            {
                sb.AppendLine(string.Join(sep,
                    r.Id,
                    Csv(r.Category, sep),
                    Csv(r.Name, sep),
                    r.Skipped ? "skipped" : Number(r.VolumeInternal, volumeUnit),
                    r.MassKnown ? Number(r.MassInternal, massUnit) : ""));
            }

            sb.AppendLine();
            sb.AppendLine(string.Join(sep, "Total volume", "", "", Number(volumeInternal, volumeUnit), ""));
            if (massValid)
            {
                sb.AppendLine(string.Join(sep, "Total mass", "", "", "", Number(massInternal, massUnit)));
            }

            XYZ source = massMode ? massCentroidInternal : centroidInternal;
            XYZ p = source - CurrentOrigin();

            sb.AppendLine(string.Join(sep, "Weighting", "", "", massMode ? "mass" : "volume", ""));
            sb.AppendLine(string.Join(sep, "Reference", "", "", selectedReference?.Label, ""));
            sb.AppendLine(string.Join(sep, "CoG X", "", "", Number(p.X, lengthUnit), ""));
            sb.AppendLine(string.Join(sep, "CoG Y", "", "", Number(p.Y, lengthUnit), ""));
            sb.AppendLine(string.Join(sep, "CoG Z", "", "", Number(p.Z, lengthUnit), ""));

            File.WriteAllText(dialog.FileName, sb.ToString(), new UTF8Encoding(true));
        }

        private static string Csv(string value, string sep)
        {
            value ??= string.Empty;
            return value.Contains(sep) || value.Contains("\"")
                ? "\"" + value.Replace("\"", "\"\"") + "\""
                : value;
        }

        private string Number(double internalValue, ForgeTypeId unit)
        {
            double display = unit != null
                ? UnitUtils.ConvertFromInternalUnits(internalValue, unit)
                : internalValue;

            return display.ToString("0.######", CultureInfo.CurrentCulture);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
