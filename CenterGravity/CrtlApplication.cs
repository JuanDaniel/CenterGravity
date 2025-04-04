﻿using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BBI.JD.Forms;
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
                CrtlApplication.thisApp.ShowForm(commandData.Application);
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
        internal static CrtlApplication thisApp = null;
        private CenterGravityForm form;

        public Result OnStartup(UIControlledApplication application)
        {
            form = null;
            thisApp = this;

            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            string folder = new FileInfo(assemblyPath).Directory.FullName;

            // Create a customm ribbon tab
            string tabName = "JDS";
            Autodesk.Windows.RibbonTab tab = CreateRibbonTab(application, tabName);

            // Add new ribbon panel
            string panelName = "Tools";
            RibbonPanel ribbonPanel = CreateRibbonPanel(application, tab, panelName);

            // Create a push button in the ribbon panel
            PushButton pushButton = ribbonPanel.AddItem(new PushButtonData(
                "CenterGravity", "Center Gravity",
                assemblyPath, "BBI.JD.Command")) as PushButton;
            
            // Set tooltip info
            pushButton.ToolTip = "Represents Center Gravity point for model elements.";

            // Set the large image shown on button
            Uri uriImage = new(Path.Combine(folder, "Resources/icon_32x32.png"));
            pushButton.LargeImage = new BitmapImage(uriImage);

            ContextualHelp help = new(ContextualHelpType.ChmFile, Path.Combine(folder, "Resources/help.chm"));
            pushButton.SetContextualHelp(help);

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            if (form != null && form.Visible)
            {
                form.Close();
            }

            return Result.Succeeded;
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

        public void ShowForm(UIApplication application)
        {
            if (form == null || form.IsDisposed)
            {
                // A new handler to handle request posting by the dialog
                RequestHandler handler = new RequestHandler();

                // External Event for the dialog to use (to post requests)
                ExternalEvent exEvent = ExternalEvent.Create(handler);

                form = new CenterGravityForm(exEvent, handler, application);
                form.Show();
            }
        }

        public void UpdateFormValues()
        {
            form?.UpdateValues();
        }

        public void UpdateFormCentroidValues()
        {
            form?.UpdateCentroidValues();
        }

        public void ShowFormMessageError(Exception ex)
        {
            form?.ShowMessageError(ex);
        }
    }
}