using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BBI.JD.UI;
using System;
using System.IO;
using System.Linq;
using System.Reflection;
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

        private static CenterGravityControl control;
        private static RequestHandler handler;
        private static ExternalEvent externalEvent;
        private static UIApplication uiApplication;

        internal static CenterGravityControl Control => control;

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

            // The WPF control is created once Revit is fully initialized (below),
            // but the pane must be registered here, in OnStartup.
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
            uiApplication = new UIApplication(sender as Application);

            handler = new RequestHandler();
            externalEvent = ExternalEvent.Create(handler);
            control = new CenterGravityControl(handler, externalEvent);

            uiApplication.SelectionChanged += OnSelectionChanged;
        }

        private static void OnSelectionChanged(object sender, Autodesk.Revit.UI.Events.SelectionChangedEventArgs e)
        {
            control?.OnRevitSelectionChanged();
        }

        internal static void RefreshPane()
        {
            control?.RefreshFromHandler();
        }

        internal static void ShowError(Exception ex)
        {
            control?.ShowError(ex);
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
