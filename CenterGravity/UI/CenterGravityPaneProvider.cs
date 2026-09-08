using Autodesk.Revit.UI;

namespace BBI.JD.UI
{
    /// <summary>Hands the WPF control to Revit when the dockable pane is first shown.</summary>
    public class CenterGravityPaneProvider : IDockablePaneProvider
    {
        public void SetupDockablePane(DockablePaneProviderData data)
        {
            data.FrameworkElement = CrtlApplication.Control;
            data.VisibleByDefault = false;

            data.InitialState = new DockablePaneState
            {
                DockPosition = DockPosition.Right
            };
        }
    }
}
