using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BBI.JD.UI;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;

namespace BBI.JD
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                // Fallback in case ApplicationInitialized never fired (or failed) for
                // this session - harmless / idempotent if it already did.
                CrtlApplication.EnsureSelectionTracking(commandData.Application);

                DockablePane pane = commandData.Application.GetDockablePane(CrtlApplication.PaneId);

                if (pane == null)
                {
                    message = "Center Gravity pane is not registered.";
                    return Result.Failed;
                }

                if (pane.IsShown())
                {
                    pane.Hide();
                }
                else
                {
                    pane.Show();
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

            return Result.Succeeded;
        }
    }

    public class CrtlApplication : IExternalApplication
    {
        public static readonly DockablePaneId PaneId =
            new DockablePaneId(new Guid("e6f1c2a3-9b84-4d5e-a7c6-1f2b3c4d5e6f"));

        // paneContent is what Revit actually displays: the real control, or - if
        // building it failed for any reason - a plain-code fallback showing the
        // exception, so the pane is never silently blank.
        private static FrameworkElement paneContent;
        private static CenterGravityControl centerGravityControl;
        private static RequestHandler handler;
        private static ExternalEvent externalEvent;
        private static UIApplication uiApplication;

        internal static FrameworkElement Control => paneContent;

        public Result OnStartup(UIControlledApplication application)
        {
            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            string folder = new FileInfo(assemblyPath).Directory.FullName;

            // Ribbon: JDS > Tools > Center Gravity
            string tabName = "JDS";
            Autodesk.Windows.RibbonTab tab = CreateRibbonTab(application, tabName);
            RibbonPanel ribbonPanel = CreateRibbonPanel(application, tab, "Tools");

            PushButton pushButton = ribbonPanel.AddItem(new PushButtonData(
                "CenterGravity", "Center Gravity",
                assemblyPath, "BBI.JD.Command")) as PushButton;

            pushButton.ToolTip = "Represents Center Gravity point for model elements.";
            pushButton.LargeImage = new BitmapImage(new Uri(Path.Combine(folder, "Resources/icon_32x32.png")));
            pushButton.SetContextualHelp(new ContextualHelp(ContextualHelpType.ChmFile, Path.Combine(folder, "Resources/help.chm")));

            // Build the pane content here, in OnStartup. This is the one point that
            // is guaranteed to run, for every Revit session, before Revit can ever
            // call SetupDockablePane - including when Revit auto-restores a pane
            // that was left open at the end of the previous session, which happens
            // during startup, before any command (and before ApplicationInitialized,
            // in some builds) has a chance to run.
            try
            {
                handler = new RequestHandler();
                externalEvent = ExternalEvent.Create(handler);
                centerGravityControl = new CenterGravityControl(handler, externalEvent);
                paneContent = centerGravityControl;
            }
            catch (Exception ex)
            {
                paneContent = BuildErrorPane(ex);
            }

            application.RegisterDockablePane(PaneId, "Center Gravity", new CenterGravityPaneProvider());

            application.ControlledApplication.ApplicationInitialized += OnApplicationInitialized;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            if (uiApplication != null)
            {
                try { uiApplication.SelectionChanged -= OnSelectionChanged; }
                catch (Exception) { /* ignore */ }
            }

            return Result.Succeeded;
        }

        private void OnApplicationInitialized(object sender, Autodesk.Revit.DB.Events.ApplicationInitializedEventArgs e)
        {
            try
            {
                EnsureSelectionTracking(new UIApplication(sender as Autodesk.Revit.ApplicationServices.Application));
            }
            catch (Exception)
            {
                // Command.Execute will retry with a known-good UIApplication.
            }
        }

        /// <summary>Idempotent - wires UIApplication.SelectionChanged exactly once.</summary>
        internal static void EnsureSelectionTracking(UIApplication application)
        {
            if (uiApplication != null || application == null)
            {
                return;
            }

            uiApplication = application;
            uiApplication.SelectionChanged += OnSelectionChanged;
        }

        private static void OnSelectionChanged(object sender, Autodesk.Revit.UI.Events.SelectionChangedEventArgs e)
        {
            centerGravityControl?.OnRevitSelectionChanged();
        }

        internal static void RefreshPane()
        {
            centerGravityControl?.RefreshFromHandler();
        }

        internal static void ShowError(Exception ex)
        {
            centerGravityControl?.ShowError(ex);
        }

        /// <summary>Pure code, no XAML - so it can render even if the real control's resources failed to load.</summary>
        private static FrameworkElement BuildErrorPane(Exception ex)
        {
            StackPanel panel = new() { Margin = new Thickness(10) };

            panel.Children.Add(new TextBlock
            {
                Text = "Center Gravity failed to start.",
                FontWeight = FontWeights.Bold,
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 8)
            });

            panel.Children.Add(new TextBlock
            {
                Text = "Please report this message to the developer:",
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 4)
            });

            panel.Children.Add(new System.Windows.Controls.TextBox
            {
                Text = ex.ToString(),
                TextWrapping = TextWrapping.Wrap,
                IsReadOnly = true,
                AcceptsReturn = true,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                MaxHeight = 400
            });

            return panel;
        }

        private Autodesk.Windows.RibbonTab CreateRibbonTab(UIControlledApplication application, string tabName)
        {
            Autodesk.Windows.RibbonTab tab = Autodesk.Windows.ComponentManager.Ribbon.Tabs.FirstOrDefault(x => x.Id == tabName);

            if (tab == null)
            {
                application.CreateRibbonTab(tabName);
                tab = Autodesk.Windows.ComponentManager.Ribbon.Tabs.FirstOrDefault(x => x.Id == tabName);
            }

            return tab;
        }

        private RibbonPanel CreateRibbonPanel(UIControlledApplication application, Autodesk.Windows.RibbonTab tab, string panelName)
        {
            RibbonPanel panel = application.GetRibbonPanels(tab.Name).FirstOrDefault(x => x.Name == panelName);

            panel ??= application.CreateRibbonPanel(tab.Name, panelName);

            return panel;
        }
    }
}
