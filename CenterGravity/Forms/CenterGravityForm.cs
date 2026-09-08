using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using System;
using System.Reflection;
using System.Windows.Forms;

namespace BBI.JD.Forms
{
    public partial class CenterGravityForm : System.Windows.Forms.Form
    {
        private readonly RequestHandler handler;
        private readonly ExternalEvent exEvent;
        private readonly UIApplication application;

        public CenterGravityForm(ExternalEvent exEvent, RequestHandler handler, UIApplication application)
        {
            InitializeComponent();

            this.exEvent = exEvent;
            this.handler = handler;
            this.application = application;
        }

        private void CenterGravity_Load(object sender, EventArgs e)
        {
            RegisterEvent();

            // Compute straight away from whatever is already selected in Revit,
            // so the tool no longer needs to be opened *before* selecting.
            MakeRequest(RequestId.Select);
        }

        private void CenterGravity_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (handler.Elements != null && handler.Elements.Count > 0)
            {
                DialogResult dialog = MessageBox.Show("Do you want to delete the center gravity drawn?", "Delete Center Gravity point", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dialog == DialogResult.Yes)
                {
                    MakeRequest(RequestId.RemoveCenterGravity);
                }
            }

            RegisterEvent(false);
        }

        private void Btn_Previous_Click(object sender, EventArgs e)
        {
            if (handler.Index > 0)
            {
                handler.ChangeIndex(-1);

                MakeRequest(RequestId.Update);
            }
            else
            {
                btn_Previous.Enabled = false;
            }

            if (handler.Index < handler.Elements.Count - 1)
            {
                btn_Next.Enabled = true;
            }
        }

        private void Btn_Next_Click(object sender, EventArgs e)
        {
            if (handler.Index < handler.Elements.Count - 1)
            {
                handler.ChangeIndex(1);

                btn_Previous.Enabled = true;

                MakeRequest(RequestId.Update);
            }
            else
            {
                btn_Next.Enabled = false;
            }
        }

        private void Btn_Clear_Click(object sender, EventArgs e)
        {
            MakeRequest(RequestId.RemoveCenterGravity);
            Clear();
        }

        private void Btn_Close_Click(object sender, EventArgs e)
        {
            Close();
        }

        private static string GetTiTleForm()
        {
            Version version = Assembly.GetExecutingAssembly().GetName().Version;

            return string.Format("{0} ({1}.{2}.{3}.{4})", "Center Gravity", version.Major, version.Minor, version.Build, version.Revision);
        }

        private void MakeRequest(RequestId request)
        {
            handler.Request.Make(request);
            exEvent.Raise();
        }

        private void RegisterEvent(bool register = true)
        {
            // Revit >= 2024 exposes a first-class selection-changed event, which
            // replaces the old trick of watching the "Modify" ribbon tab title.
            if (register)
            {
                application.SelectionChanged += OnRevitSelectionChanged;
            }
            else
            {
                application.SelectionChanged -= OnRevitSelectionChanged;
            }
        }

        private void OnRevitSelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            btn_Previous.Enabled = false;
            btn_Next.Enabled = false;

            MakeRequest(RequestId.Select);
        }

        private void Clear()
        {
            lbl_Index.Text = string.Empty;
            txt_ID.Clear();
            txt_Name.Clear();
            txt_Family.Clear();
            txt_Volume.Clear();
            txt_X.Clear();
            txt_Y.Clear();
            txt_Z.Clear();
            txt_XYZ.Clear();
        }

        public void UpdateValues()
        {
            if (handler.Elements == null || handler.Elements.Count == 0)
            {
                lbl_Index.Text = string.Empty;
                btn_Next.Enabled = false;
                btn_Previous.Enabled = false;
                Clear();

                return;
            }

            btn_Next.Enabled = handler.Elements.Count > 1;

            lbl_Index.Text = string.Format("{0} / {1}", handler.Index + 1, handler.Elements.Count);

            if (handler.Index > -1)
            {
                Element element = handler.Elements[handler.Index];

                txt_ID.Text = element.Id.ToString();
                txt_Name.Text = element.Name;
                txt_Family.Text = element.GetType().Name;
            }
            else
            {
                // Clear all values
                Clear();
            }
        }

        public void UpdateCentroidValues()
        {
            if (handler.CV == null || !handler.CV.IsValid)
            {
                // Selection has no usable solid geometry.
                txt_Volume.Text = "-";
                txt_X.Text = "-";
                txt_Y.Text = "-";
                txt_Z.Text = "-";
                txt_XYZ.Text = "-";

                return;
            }

            Units units = application.ActiveUIDocument.Document.GetUnits();

            FormatOptions fo_volume;
            FormatOptions fo_length;

            XYZ vector;

            fo_volume = units.GetFormatOptions(SpecTypeId.Volume);
            fo_length = units.GetFormatOptions(SpecTypeId.Length);

            // Convert to current display volume units
            txt_Volume.Text = string.Format("{0:0.000}", UnitUtils.ConvertFromInternalUnits(handler.CV.Volume, fo_volume.GetUnitTypeId()));

            // Convert to current display length units
            vector = new XYZ(
                UnitUtils.ConvertFromInternalUnits(handler.CV.Centroid.X, fo_length.GetUnitTypeId()),
                UnitUtils.ConvertFromInternalUnits(handler.CV.Centroid.Y, fo_length.GetUnitTypeId()),
                UnitUtils.ConvertFromInternalUnits(handler.CV.Centroid.Z, fo_length.GetUnitTypeId())
            );

            txt_X.Text = string.Format("{0:0.000}", vector.X);
            txt_Y.Text = string.Format("{0:0.000}", vector.Y);
            txt_Z.Text = string.Format("{0:0.000}", vector.Z);

            txt_XYZ.Text = handler.CV.XYZToString(fo_length);
        }

        public void ShowMessageError(Exception ex)
        {
            MessageBox.Show(ex.Message, "Center Gravity error exception", MessageBoxButtons.OK, MessageBoxIcon.Error);

            handler.ClearElements();

            Close();
        }

        private void CenterGravityForm_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }
    }
}
