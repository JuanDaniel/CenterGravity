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
        private readonly List<XYZ> liftPoints = new();
        private RiggingResult rigging;

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

        /// <summary>Origin (internal coords) the pane is reporting against; written onto the marker.</summary>
        public XYZ ReferenceOrigin { get; set; } = XYZ.Zero;

        public string ReferenceLabel { get; set; }
        public string LiftName { get; set; }

        /// <summary>How many lift points the next pick should collect (2 or 4).</summary>
        public int LiftPointCount { get; set; } = 2;

        /// <summary>Hook height above the centre of gravity (internal length), or null.</summary>
        public double? HookHeightInternal { get; set; }

        public IReadOnlyList<XYZ> LiftPoints => liftPoints;
        public RiggingResult Rigging => rigging;

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
                    case RequestId.CreateSchedule:
                        CreateSchedule(application);
                        break;
                    case RequestId.PickLiftPoints:
                        PickLiftPoints(application);
                        break;
                    case RequestId.ClearLiftPoints:
                        liftPoints.Clear();
                        rigging = null;
                        CrtlApplication.RefreshPane();
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

            // Lift points belong to the previous centre of gravity.
            liftPoints.Clear();
            rigging = null;

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
            XYZ reported = location - (ReferenceOrigin ?? XYZ.Zero);

            EnsureFamilyLoaded(application);

            FamilySymbol familySymbol = family?
                .GetFamilySymbolIds()
                .Select(id => document.GetElement(id) as FamilySymbol)
                .FirstOrDefault(fs => fs != null && fs.FamilyName == FamilyName);

            if (familySymbol == null)
            {
                return;
            }

            // Best effort - a marker without schedulable parameters is still useful.
            try
            {
                CgSharedParameters.EnsureBound(document, familySymbol.Category);
            }
            catch (Exception)
            {
                // carry on without shared parameters
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

            string coordinateText = FormatPoint(reported);

            Parameter legacy = familyInstance.LookupParameter(CoordinateParameterName);
            legacy?.Set(coordinateText);

            WriteMarkerParameters(familyInstance, reported, coordinateText, useMass);

            markerIds.Add(familyInstance.Id);

            transaction.Commit();

            CrtlApplication.RefreshPane();
        }

        private void WriteMarkerParameters(FamilyInstance fi, XYZ reported, string coordinateText, bool useMass)
        {
            SetParam(fi, CgSharedParameters.Marker, 1);
            SetParam(fi, CgSharedParameters.LiftName, LiftName ?? string.Empty);
            SetParam(fi, CgSharedParameters.Weighting, useMass ? "mass" : "volume");
            SetParam(fi, CgSharedParameters.Reference, string.IsNullOrEmpty(ReferenceLabel) ? "Internal origin" : ReferenceLabel);
            SetParam(fi, CgSharedParameters.Coordinates, coordinateText);
            SetParam(fi, CgSharedParameters.X, reported.X);
            SetParam(fi, CgSharedParameters.Y, reported.Y);
            SetParam(fi, CgSharedParameters.Z, reported.Z);
            SetParam(fi, CgSharedParameters.Date, DateTime.Now.ToString("yyyy-MM-dd HH:mm"));

            if (cv != null && cv.IsValid)
            {
                SetParam(fi, CgSharedParameters.Volume, cv.Volume);
            }

            if (useMass && cv != null && cv.MassIsValid)
            {
                SetParam(fi, CgSharedParameters.Weight, cv.Mass);
            }
        }

        private static void SetParam(FamilyInstance fi, CgSharedParameters.ParamDef def, double value)
        {
            Parameter p = fi.get_Parameter(def.Guid);
            if (p != null && !p.IsReadOnly)
            {
                p.Set(value);
            }
        }

        private static void SetParam(FamilyInstance fi, CgSharedParameters.ParamDef def, int value)
        {
            Parameter p = fi.get_Parameter(def.Guid);
            if (p != null && !p.IsReadOnly)
            {
                p.Set(value);
            }
        }

        private static void SetParam(FamilyInstance fi, CgSharedParameters.ParamDef def, string value)
        {
            Parameter p = fi.get_Parameter(def.Guid);
            if (p != null && !p.IsReadOnly)
            {
                p.Set(value ?? string.Empty);
            }
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

        private void CreateSchedule(UIApplication application)
        {
            UIDocument uiDoc = application.ActiveUIDocument;
            Document document = uiDoc.Document;

            Category category = GetMarkerCategory(document);
            if (category == null)
            {
                return;
            }

            try
            {
                CgSharedParameters.EnsureBound(document, category);
            }
            catch (Exception)
            {
                // continue - schedule can still be created with whatever is bound
            }

            Dictionary<string, ElementId> paramIds = new();
            foreach (CgSharedParameters.ParamDef d in CgSharedParameters.All)
            {
                SharedParameterElement spe = SharedParameterElement.Lookup(document, d.Guid);
                if (spe != null)
                {
                    paramIds[d.Name] = spe.Id;
                }
            }

            string[] columns =
            {
                "CG_Marker", "CG_LiftName", "CG_Weight", "CG_Volume",
                "CG_X", "CG_Y", "CG_Z", "CG_Reference", "CG_Weighting", "CG_Date"
            };

            using Transaction transaction = new(document);
            transaction.Start("Create Center of Gravity schedule");

            ViewSchedule schedule = ViewSchedule.CreateSchedule(document, category.Id);
            schedule.Name = UniqueViewName(document, "Center of Gravity");

            ScheduleDefinition definition = schedule.Definition;
            IList<SchedulableField> schedulable = definition.GetSchedulableFields();

            ScheduleField markerField = null;

            foreach (string name in columns)
            {
                if (!paramIds.TryGetValue(name, out ElementId pid))
                {
                    continue;
                }

                SchedulableField sf = schedulable.FirstOrDefault(x => x.ParameterId == pid);
                if (sf == null)
                {
                    continue;
                }

                ScheduleField field = definition.AddField(sf);

                if (name == "CG_Marker")
                {
                    markerField = field;
                    field.IsHidden = true;
                }
            }

            if (markerField != null)
            {
                definition.AddFilter(new ScheduleFilter(markerField.FieldId, ScheduleFilterType.Equal, 1));
            }

            transaction.Commit();

            uiDoc.RequestViewChange(schedule);
        }

        private void PickLiftPoints(UIApplication application)
        {
            UIDocument uiDoc = application.ActiveUIDocument;

            if (cv == null || !cv.IsValid)
            {
                return;
            }

            int count = LiftPointCount == 4 ? 4 : 2;
            liftPoints.Clear();

            try
            {
                for (int i = 0; i < count; i++)
                {
                    XYZ p = uiDoc.Selection.PickPoint(
                        string.Format("Center Gravity: pick lift point {0} of {1} (Esc to stop)", i + 1, count));
                    liftPoints.Add(p);
                }
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                // keep whatever was picked so far
            }

            ComputeRigging();
            CrtlApplication.RefreshPane();
        }

        private void ComputeRigging()
        {
            rigging = null;

            if (cv == null || !cv.IsValid || liftPoints.Count < 2)
            {
                return;
            }

            bool massMode = PlaceAtMass && cv.MassIsValid;
            XYZ cog = massMode ? cv.MassCentroid : cv.Centroid;

            double weight = 0.0;
            if (cv.MassIsValid)
            {
                weight = cv.Mass;
            }
            else if (DefaultDensityInternal.HasValue && DefaultDensityInternal.Value > 0)
            {
                weight = cv.Volume * DefaultDensityInternal.Value;
            }

            rigging = RiggingCalculator.Compute(cog, weight, liftPoints, HookHeightInternal ?? 0.0);
        }

        private Category GetMarkerCategory(Document document)
        {
            if (family != null && family.IsValidObject)
            {
                foreach (ElementId id in family.GetFamilySymbolIds())
                {
                    if (document.GetElement(id) is FamilySymbol fs && fs.Category != null)
                    {
                        return fs.Category;
                    }
                }
            }

            return Category.GetCategory(document, BuiltInCategory.OST_GenericModel);
        }

        private static string UniqueViewName(Document document, string baseName)
        {
            HashSet<string> used = new(new FilteredElementCollector(document)
                .OfClass(typeof(View))
                .Cast<View>()
                .Select(v => v.Name));

            if (!used.Contains(baseName))
            {
                return baseName;
            }

            for (int i = 2; i < 1000; i++)
            {
                string candidate = $"{baseName} {i}";
                if (!used.Contains(candidate))
                {
                    return candidate;
                }
            }

            return $"{baseName} {Guid.NewGuid():N}";
        }
    }
}
