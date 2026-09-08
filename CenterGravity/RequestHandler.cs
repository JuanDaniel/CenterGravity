using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using BBI.JD.UI;
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
        private List<Element> elements = new();
        private CentroidVolume cv;
        private List<CgRow> rows = new();
        private Units units;
        private XYZ projectBasePoint = XYZ.Zero;
        private XYZ surveyPoint = XYZ.Zero;
        private int expandedContainerCount;

        public Request Request => request;

        public string GetName() => "Center Gravity";

        public CentroidVolume CV => cv;
        public IReadOnlyList<CgRow> Rows => rows;
        public Units CurrentUnits => units;
        public XYZ ProjectBasePoint => projectBasePoint;
        public XYZ SurveyPoint => surveyPoint;
        public bool HasMarkers => markerIds.Count > 0;
        public bool HasResult => cv != null && cv.IsValid;
        public int ExpandedContainerCount => expandedContainerCount;

        /// <summary>Fallback mass density (internal units) applied when a material has none. Null = no fallback.</summary>
        public double? DefaultDensityInternal { get; set; }

        /// <summary>When true, "Place marker" drops the point at the mass-weighted centre of gravity.</summary>
        public bool PlaceAtMass { get; set; }

        public void Execute(UIApplication application)
        {
            try
            {
                switch (Request.Take())
                {
                    case RequestId.None:
                        return;
                    case RequestId.CenterGravityFamily:
                        EnsureFamilyLoaded(application);
                        break;
                    case RequestId.Select:
                        Recompute(application);
                        break;
                    case RequestId.PlaceCenterGravity:
                        PlaceMarker(application);
                        break;
                    case RequestId.RemoveCenterGravity:
                        RemoveMarkers(application);
                        break;
                }
            }
            catch (Exception ex)
            {
                CrtlApplication.ShowError(ex);
            }
        }

        private void Recompute(UIApplication application)
        {
            UIDocument uiDoc = application.ActiveUIDocument;
            Document document = uiDoc.Document;

            units = document.GetUnits();
            ReadReferencePoints(document);

            elements = new List<Element>();
            cv = null;
            rows = new List<CgRow>();
            expandedContainerCount = 0;

            ICollection<ElementId> ids = uiDoc.Selection.GetElementIds();

            if (ids.Count > 0)
            {
                elements = ExpandContainers(document, ids, out expandedContainerCount)
                    .Where(e => e.IsPhysicalElement())
                    .ToList();
            }

            if (elements.Count > 0)
            {
                IDensityProvider density = new RevitDensityProvider(document, DefaultDensityInternal);

                cv = GeometryUtils.GetCentroid(elements, new Options(), density);
                rows = BuildRows();
            }

            CrtlApplication.RefreshPane();
        }

        /// <summary>Replace groups / assemblies in the selection with their member elements (nested too).</summary>
        private static List<Element> ExpandContainers(Document document, ICollection<ElementId> ids, out int containerCount)
        {
            containerCount = 0;

            List<Element> result = new();
            HashSet<ElementId> seen = new();
            Queue<ElementId> queue = new(ids);
            int guard = 0;

            while (queue.Count > 0 && guard++ < 50000)
            {
                ElementId id = queue.Dequeue();

                if (!seen.Add(id))
                {
                    continue;
                }

                Element e = document.GetElement(id);

                if (e == null)
                {
                    continue;
                }

                if (e is Group group)
                {
                    containerCount++;
                    foreach (ElementId member in group.GetMemberIds())
                    {
                        queue.Enqueue(member);
                    }
                    continue;
                }

                if (e is AssemblyInstance assembly)
                {
                    containerCount++;
                    foreach (ElementId member in assembly.GetMemberIds())
                    {
                        queue.Enqueue(member);
                    }
                    continue;
                }

                result.Add(e);
            }

            return result;
        }

        private List<CgRow> BuildRows()
        {
            Dictionary<ElementId, ElementContribution> byId = cv.Contributions
                .GroupBy(c => c.Id)
                .ToDictionary(g => g.Key, g => g.First());

            HashSet<ElementId> skipped = new(cv.SkippedElementIds);
            ForgeTypeId massSpec = SpecTypeId.Mass;

            List<CgRow> result = new();

            foreach (Element e in elements)
            {
                bool isSkipped = skipped.Contains(e.Id) || !byId.ContainsKey(e.Id);
                ElementContribution c = isSkipped ? null : byId[e.Id];

                result.Add(new CgRow
                {
                    Id = e.Id.ToString(),
                    Category = e.Category?.Name ?? string.Empty,
                    Name = e.Name,
                    Skipped = isSkipped,
                    VolumeInternal = c?.Volume ?? 0.0,
                    Volume = isSkipped
                        ? "-"
                        : UnitFormatUtils.Format(units, SpecTypeId.Volume, c.Volume, false),
                    MassInternal = c?.Mass ?? 0.0,
                    MassKnown = c != null && c.MassKnown,
                    Mass = (c != null && c.MassKnown)
                        ? UnitFormatUtils.Format(units, massSpec, c.Mass, false)
                        : "-"
                });
            }

            return result;
        }

        private void ReadReferencePoints(Document document)
        {
            try
            {
                BasePoint pbp = BasePoint.GetProjectBasePoint(document);
                BasePoint sp = BasePoint.GetSurveyPoint(document);

                projectBasePoint = pbp?.Position ?? XYZ.Zero;
                surveyPoint = sp?.Position ?? XYZ.Zero;
            }
            catch (Exception)
            {
                projectBasePoint = XYZ.Zero;
                surveyPoint = XYZ.Zero;
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

        private void PlaceMarker(UIApplication application)
        {
            Document document = application.ActiveUIDocument.Document;

            if (cv == null || !cv.IsValid || elements.Count == 0)
            {
                return;
            }

            bool useMass = PlaceAtMass && cv.MassIsValid;
            XYZ location = useMass ? cv.MassCentroid : cv.Centroid;

            EnsureFamilyLoaded(application);

            FamilySymbol familySymbol = family?
                .GetFamilySymbolIds()
                .Select(id => document.GetElement(id) as FamilySymbol)
                .FirstOrDefault(fs => fs != null && fs.FamilyName == FamilyName);

            if (familySymbol == null)
            {
                return;
            }

            Element host = elements[0];

            using Transaction transaction = new(document);
            transaction.Start("Place Center Gravity point");

            if (!familySymbol.IsActive)
            {
                familySymbol.Activate();
            }

            Level level = host.LevelId != ElementId.InvalidElementId
                ? document.GetElement(host.LevelId) as Level
                : null;

            FamilyInstance familyInstance = level != null
                ? document.Create.NewFamilyInstance(location, familySymbol, host, level, StructuralType.NonStructural)
                : document.Create.NewFamilyInstance(location, familySymbol, host, StructuralType.NonStructural);

            Parameter coordinate = familyInstance.LookupParameter(CoordinateParameterName);
            coordinate?.Set(FormatPoint(location));

            markerIds.Add(familyInstance.Id);

            transaction.Commit();

            CrtlApplication.RefreshPane();
        }

        private string FormatPoint(XYZ p)
        {
            FormatOptions fo = units.GetFormatOptions(SpecTypeId.Length);

            return string.Format("{0}; {1}; {2}",
                UnitUtils.ConvertFromInternalUnits(p.X, fo.GetUnitTypeId()),
                UnitUtils.ConvertFromInternalUnits(p.Y, fo.GetUnitTypeId()),
                UnitUtils.ConvertFromInternalUnits(p.Z, fo.GetUnitTypeId()));
        }

        private void RemoveMarkers(UIApplication application)
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

            CrtlApplication.RefreshPane();
        }
    }
}
