using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace BBI.JD
{
    public class RequestHandler : IExternalEventHandler
    {
        private const string FamilyName = "CenterGravityFamily";
        private const string CoordinateParameterName = "CenterGravity";

        private readonly Request request = new();
        private readonly List<ElementId> markerIds = new();

        private Family family;
        private int index = -1;
        private List<Element> elements = new();
        private CentroidVolume cv;

        public Request Request
        {
            get { return request; }
        }

        public string GetName()
        {
            return "Center Gravity";
        }

        public int Index
        {
            get { return index; }
        }

        public List<Element> Elements
        {
            get { return elements; }
        }

        public CentroidVolume CV
        {
            get { return cv; }
        }

        public int ChangeIndex(int value)
        {
            index += value;

            return index;
        }

        public void ClearElements()
        {
            elements?.Clear();
        }

        public void Execute(UIApplication application)
        {
            try
            {
                switch (Request.Take())
                {
                    case RequestId.None:
                        {
                            return;
                        }
                    case RequestId.CenterGravityFamily:
                        {
                            EnsureFamilyLoaded(application);
                            break;
                        }
                    case RequestId.Select:
                        {
                            SelectionChanged(application);
                            break;
                        }
                    case RequestId.Update:
                        {
                            UpdateValues(application);
                            break;
                        }
                    case RequestId.VisualizeCenterGravity:
                        {
                            UpdateCentroidValues(application);
                            break;
                        }
                    case RequestId.RemoveCenterGravity:
                        {
                            RemoveCenterGravityPoints(application);
                            break;
                        }
                }
            }
            catch (Exception ex)
            {
                CrtlApplication.thisApp.ShowFormMessageError(ex);
            }
        }

        private void EnsureFamilyLoaded(UIApplication application)
        {
            Document document = application.ActiveUIDocument.Document;

            if (family != null && family.IsValidObject)
            {
                return;
            }

            family = new FilteredElementCollector(document)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .FirstOrDefault(f => f.Name == FamilyName);

            if (family == null)
            {
                string folder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

                using Transaction transaction = new(document);
                transaction.Start("Load CenterGravityFamily");

                document.LoadFamily(Path.Combine(folder, "Resources", FamilyName + ".rfa"), out family);

                transaction.Commit();
            }
        }

        private void SelectionChanged(UIApplication application)
        {
            UIDocument uiDoc = application.ActiveUIDocument;
            Document document = uiDoc.Document;

            // Reset INDEX and ELEMENTS
            index = -1;
            elements = new List<Element>();
            cv = null;

            ICollection<ElementId> ids = uiDoc.Selection.GetElementIds();

            if (ids.Count > 0)
            {
                elements = new FilteredElementCollector(document, ids)
                    .WhereElementIsNotElementType()
                        .Where(e => e.IsPhysicalElement())
                            .ToList();

                if (elements.Count > 0)
                {
                    index = 0;

                    UpdateCentroidValues(application);
                }
            }

            UpdateValues(application);
        }

        private void UpdateValues(UIApplication application)
        {
            CrtlApplication.thisApp.UpdateFormValues();
        }

        private void UpdateCentroidValues(UIApplication application)
        {
            VisualizeCentroid(application);

            CrtlApplication.thisApp.UpdateFormCentroidValues();
        }

        private void VisualizeCentroid(UIApplication application)
        {
            Document document = application.ActiveUIDocument.Document;

            if (elements == null || elements.Count == 0)
            {
                return;
            }

            cv = GeometryUtils.GetCentroid(elements, new Options());

            // No usable solid geometry in the selection: keep the (zeroed) result
            // so the form can report it, but do not try to place a marker.
            if (cv == null || !cv.IsValid)
            {
                return;
            }

            EnsureFamilyLoaded(application);

            if (family == null)
            {
                return;
            }

            FamilySymbol familySymbol = family.GetFamilySymbolIds()
                .Select(id => document.GetElement(id) as FamilySymbol)
                .FirstOrDefault(fs => fs != null && fs.FamilyName == FamilyName);

            if (familySymbol == null)
            {
                return;
            }

            Element host = elements[0];

            using Transaction transaction = new(document);
            transaction.Start("Put graphical Center Gravity point");

            if (!familySymbol.IsActive)
            {
                familySymbol.Activate();
            }

            Level level = host.LevelId != ElementId.InvalidElementId
                ? document.GetElement(host.LevelId) as Level
                : null;

            FamilyInstance familyInstance = level != null
                ? document.Create.NewFamilyInstance(cv.Centroid, familySymbol, host, level, StructuralType.NonStructural)
                : document.Create.NewFamilyInstance(cv.Centroid, familySymbol, host, StructuralType.NonStructural);

            Parameter coordinate = familyInstance.LookupParameter(CoordinateParameterName);
            coordinate?.Set(cv.XYZToString(
                document.GetUnits().GetFormatOptions(SpecTypeId.Length)
            ));

            markerIds.Add(familyInstance.Id);

            transaction.Commit();
        }

        private void RemoveCenterGravityPoints(UIApplication application)
        {
            Document document = application.ActiveUIDocument.Document;

            List<ElementId> toDelete = markerIds
                .Where(id => document.GetElement(id) != null)
                .ToList();

            if (toDelete.Count > 0)
            {
                using Transaction transaction = new(document);
                transaction.Start("Remove Center Gravity points");

                document.Delete(toDelete);

                transaction.Commit();
            }

            markerIds.Clear();
            elements = new List<Element>();
            index = -1;
            cv = null;
        }
    }
}
