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

    public class CenterGravityViewModel : INotifyPropertyChanged
    {
        private Units units;
        private ForgeTypeId lengthUnit;
        private ForgeTypeId volumeUnit;

        private XYZ centroidInternal = XYZ.Zero;
        private XYZ projectBasePoint = XYZ.Zero;
        private XYZ surveyPoint = XYZ.Zero;
        private double volumeInternal;
        private bool hasResult;
        private bool hasMarkers;
        private int skippedCount;

        public CenterGravityViewModel()
        {
            ReferenceOptions = new[]
            {
                new ReferenceOption { Mode = ReferenceMode.InternalOrigin,   Label = "Internal origin" },
                new ReferenceOption { Mode = ReferenceMode.ProjectBasePoint, Label = "Project base point" },
                new ReferenceOption { Mode = ReferenceMode.SurveyPoint,      Label = "Survey point" },
            };
            selectedReference = ReferenceOptions[0];

            PlaceCommand = new RelayCommand(() => PlaceRequested?.Invoke(), () => hasResult);
            ClearCommand = new RelayCommand(() => ClearRequested?.Invoke(), () => hasMarkers);
            CopyCommand = new RelayCommand(CopyCoordinates, () => hasResult);
            ExportCsvCommand = new RelayCommand(ExportCsv, () => hasResult);
        }

        /// <summary>Raised when the user asks to drop the centre-of-gravity marker into the model.</summary>
        public Action PlaceRequested;

        /// <summary>Raised when the user asks to remove the drawn centre-of-gravity marker(s).</summary>
        public Action ClearRequested;

        public ObservableCollection<CgRow> Rows { get; } = new ObservableCollection<CgRow>();

        public ReferenceOption[] ReferenceOptions { get; }

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
        public string X { get; private set; } = "-";
        public string Y { get; private set; } = "-";
        public string Z { get; private set; } = "-";
        public string Xyz { get; private set; } = "-";

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
                    text += string.Format(" - {0} skipped (no solid geometry)", skippedCount);
                }

                return text;
            }
        }

        /// <summary>Feed a fresh result in (called on the UI thread by the pane).</summary>
        public void SetResult(Units units, CentroidVolume cv, IEnumerable<CgRow> rows, XYZ projectBasePoint, XYZ surveyPoint)
        {
            this.units = units;
            lengthUnit = units?.GetFormatOptions(SpecTypeId.Length)?.GetUnitTypeId();
            volumeUnit = units?.GetFormatOptions(SpecTypeId.Volume)?.GetUnitTypeId();

            this.projectBasePoint = projectBasePoint ?? XYZ.Zero;
            this.surveyPoint = surveyPoint ?? XYZ.Zero;

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

            OnPropertyChanged(nameof(HasResult));
            OnPropertyChanged(nameof(Summary));
            RefreshCoordinates();

            CommandManager.InvalidateRequerySuggested();
        }

        public void Reset()
        {
            Rows.Clear();
            skippedCount = 0;
            hasResult = false;
            centroidInternal = XYZ.Zero;
            volumeInternal = 0.0;

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
            if (!hasResult)
            {
                Volume = X = Y = Z = Xyz = "-";
            }
            else
            {
                XYZ p = centroidInternal - CurrentOrigin();

                Volume = FormatLength(volumeInternal, volumeUnit, SpecTypeId.Volume);
                X = FormatLength(p.X, lengthUnit, SpecTypeId.Length);
                Y = FormatLength(p.Y, lengthUnit, SpecTypeId.Length);
                Z = FormatLength(p.Z, lengthUnit, SpecTypeId.Length);
                Xyz = string.Format("{0}; {1}; {2}", X, Y, Z);
            }

            OnPropertyChanged(nameof(Volume));
            OnPropertyChanged(nameof(X));
            OnPropertyChanged(nameof(Y));
            OnPropertyChanged(nameof(Z));
            OnPropertyChanged(nameof(Xyz));
        }

        private string FormatLength(double internalValue, ForgeTypeId unit, ForgeTypeId spec)
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

            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Join(sep, "Id", "Category", "Name", "Volume"));

            foreach (CgRow r in Rows)
            {
                sb.AppendLine(string.Join(sep,
                    r.Id,
                    Csv(r.Category, sep),
                    Csv(r.Name, sep),
                    r.Skipped ? "skipped" : Number(r.VolumeInternal, volumeUnit)));
            }

            sb.AppendLine();
            sb.AppendLine(string.Join(sep, "Total volume", "", "", Number(volumeInternal, volumeUnit)));

            XYZ p = centroidInternal - CurrentOrigin();
            sb.AppendLine(string.Join(sep, "Reference", "", "", selectedReference?.Label));
            sb.AppendLine(string.Join(sep, "CoG X", "", "", Number(p.X, lengthUnit)));
            sb.AppendLine(string.Join(sep, "CoG Y", "", "", Number(p.Y, lengthUnit)));
            sb.AppendLine(string.Join(sep, "CoG Z", "", "", Number(p.Z, lengthUnit)));

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
