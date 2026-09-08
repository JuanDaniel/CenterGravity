using Autodesk.Revit.UI;
using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace BBI.JD.UI
{
    public partial class CenterGravityControl : UserControl
    {
        private readonly RequestHandler handler;
        private readonly ExternalEvent exEvent;
        private readonly CenterGravityViewModel vm;

        public CenterGravityControl(RequestHandler handler, ExternalEvent exEvent)
        {
            InitializeComponent();

            this.handler = handler;
            this.exEvent = exEvent;

            vm = new CenterGravityViewModel();
            vm.PlaceRequested = () =>
            {
                this.handler.PlaceAtMass = vm.WeightByMass;
                this.handler.ReferenceOrigin = vm.CurrentReferenceOriginInternal;
                this.handler.ReferenceLabel = vm.CurrentReferenceLabel;
                this.handler.LiftName = vm.LiftName;
                MakeRequest(RequestId.PlaceCenterGravity);
            };
            vm.ClearRequested = () => MakeRequest(RequestId.RemoveCenterGravity);
            vm.CreateScheduleRequested = () => MakeRequest(RequestId.CreateSchedule);
            vm.DensityChanged = () =>
            {
                this.handler.DefaultDensityInternal = vm.DefaultDensityInternal;
                MakeRequest(RequestId.Select);
            };
            vm.WeightModeChanged = () => this.handler.PlaceAtMass = vm.WeightByMass;

            DataContext = vm;

            ApplyTheme();
            Loaded += (s, e) => MakeRequest(RequestId.Select);
        }

        /// <summary>Called from the Revit UIApplication.SelectionChanged event.</summary>
        public void OnRevitSelectionChanged()
        {
            MakeRequest(RequestId.Select);
        }

        /// <summary>Called by the request handler once a calculation (or marker change) is done.</summary>
        public void RefreshFromHandler()
        {
            void Apply()
            {
                vm.HasMarkers = handler.HasMarkers;
                vm.SetResult(
                    handler.CurrentUnits,
                    handler.CV,
                    handler.Rows,
                    handler.ProjectBasePoint,
                    handler.SurveyPoint,
                    handler.ExpandedContainerCount);
            }

            if (Dispatcher.CheckAccess())
            {
                Apply();
            }
            else
            {
                Dispatcher.Invoke(Apply);
            }
        }

        public void ShowError(Exception ex)
        {
            void Show() => MessageBox.Show(ex.Message, "Center Gravity", MessageBoxButton.OK, MessageBoxImage.Error);

            if (Dispatcher.CheckAccess())
            {
                Show();
            }
            else
            {
                Dispatcher.Invoke(Show);
            }
        }

        private void MakeRequest(RequestId request)
        {
            handler.Request.Make(request);
            exEvent.Raise();
        }

        private void ApplyTheme()
        {
            // Follow the Revit UI theme (available since Revit 2024).
            bool dark;
            try
            {
                dark = UIThemeManager.CurrentTheme == UITheme.Dark;
            }
            catch (Exception)
            {
                dark = false;
            }

            if (dark)
            {
                Background = new SolidColorBrush(Color.FromRgb(0x2E, 0x2E, 0x2E));
                Foreground = new SolidColorBrush(Color.FromRgb(0xE6, 0xE6, 0xE6));
            }
            else
            {
                Background = new SolidColorBrush(Color.FromRgb(0xF5, 0xF5, 0xF5));
                Foreground = new SolidColorBrush(Color.FromRgb(0x1E, 0x1E, 0x1E));
            }
        }
    }
}
