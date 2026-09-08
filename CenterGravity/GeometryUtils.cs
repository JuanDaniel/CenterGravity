using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;

namespace BBI.JD
{
    /// <summary>Per-element contribution to a combined centre of gravity.</summary>
    public class ElementContribution
    {
        public ElementId Id { get; set; }
        public double Volume { get; set; }
        public XYZ Centroid { get; set; }

        /// <summary>Mass in Revit internal units; 0 when no density was available.</summary>
        public double Mass { get; set; }

        /// <summary>Effective mass density used (internal units); 0 when unknown.</summary>
        public double Density { get; set; }

        /// <summary>True when a real mass (from material density) was found for this element.</summary>
        public bool MassKnown { get; set; }
    }

    public class CentroidVolume
    {
        public CentroidVolume()
        {
            Centroid = XYZ.Zero;
            Volume = 0.0;
            MassCentroid = XYZ.Zero;
            Mass = 0.0;
            SkippedElementIds = new List<ElementId>();
            NoDensityElementIds = new List<ElementId>();
            Contributions = new List<ElementContribution>();
        }

        public XYZ Centroid { get; set; }
        public double Volume { get; set; }

        /// <summary>Mass-weighted centre of gravity (internal units). Zero when no mass is known.</summary>
        public XYZ MassCentroid { get; set; }

        /// <summary>Total mass of the elements that had a known density (internal units).</summary>
        public double Mass { get; set; }

        /// <summary>True when every contributing element had a known mass, so MassCentroid is exact.</summary>
        public bool MassComplete { get; set; }

        /// <summary>Elements that carried no usable solid geometry and were left out of the calculation.</summary>
        public List<ElementId> SkippedElementIds { get; }

        /// <summary>Elements that contributed volume but had no density (excluded from the mass result).</summary>
        public List<ElementId> NoDensityElementIds { get; }

        /// <summary>Per-element volume / centroid / mass, in the order the elements were supplied.</summary>
        public List<ElementContribution> Contributions { get; }

        /// <summary>True when the result is a real, finite centroid backed by a non-zero volume.</summary>
        public bool IsValid
        {
            get
            {
                return Centroid != null
                    && Math.Abs(Volume) > GeometryUtils.VolumeTolerance
                    && !double.IsNaN(Centroid.X) && !double.IsNaN(Centroid.Y) && !double.IsNaN(Centroid.Z)
                    && !double.IsInfinity(Centroid.X) && !double.IsInfinity(Centroid.Y) && !double.IsInfinity(Centroid.Z);
            }
        }

        /// <summary>True when the mass-weighted centre of gravity is finite and backed by a non-zero mass.</summary>
        public bool MassIsValid
        {
            get
            {
                return MassCentroid != null
                    && Math.Abs(Mass) > GeometryUtils.MassTolerance
                    && !double.IsNaN(MassCentroid.X) && !double.IsNaN(MassCentroid.Y) && !double.IsNaN(MassCentroid.Z)
                    && !double.IsInfinity(MassCentroid.X) && !double.IsInfinity(MassCentroid.Y) && !double.IsInfinity(MassCentroid.Z);
            }
        }

        public string XYZToString(FormatOptions fo)
        {
            // Convert to current display length units
            return string.Format("{0}; {1}; {2}",
                UnitUtils.ConvertFromInternalUnits(Centroid.X, fo.GetUnitTypeId()),
                UnitUtils.ConvertFromInternalUnits(Centroid.Y, fo.GetUnitTypeId()),
                UnitUtils.ConvertFromInternalUnits(Centroid.Z, fo.GetUnitTypeId())
            );
        }
    }

    static class GeometryUtils
    {
        // Volumes are handled in internal units (cubic feet); anything below this
        // is treated as "no volume" to avoid dividing by (almost) zero.
        internal const double VolumeTolerance = 1e-9;

        // Same idea for mass (internal units).
        internal const double MassTolerance = 1e-6;

        public static CentroidVolume GetCentroid(Solid solid)
        {
            CentroidVolume cv = new();
            double v;
            XYZ v0, v1, v2;

            SolidOrShellTessellationControls controls = new()
            {
                LevelOfDetail = 0
            };

            TriangulatedSolidOrShell triangulation = null;

            try
            {
                triangulation = SolidUtils.TessellateSolidOrShell(
                    solid, controls);
            }
            catch (Autodesk.Revit.Exceptions.InvalidOperationException)
            {
                return null;
            }

            int n = triangulation.ShellComponentCount;

            for (int i = 0; i < n; ++i)
            {
                TriangulatedShellComponent component = triangulation.GetShellComponent(i);

                int m = component.TriangleCount;

                for (int j = 0; j < m; ++j)
                {
                    TriangleInShellComponent t = component.GetTriangle(j);

                    v0 = component.GetVertex(t.VertexIndex0);
                    v1 = component.GetVertex(t.VertexIndex1);
                    v2 = component.GetVertex(t.VertexIndex2);

                    v = v0.X * (v1.Y * v2.Z - v2.Y * v1.Z)
                      + v0.Y * (v1.Z * v2.X - v2.Z * v1.X)
                      + v0.Z * (v1.X * v2.Y - v2.X * v1.Y);

                    cv.Centroid += v * (v0 + v1 + v2);
                    cv.Volume += v;
                }
            }

            // Degenerate / empty solid: no meaningful centroid.
            if (Math.Abs(cv.Volume) < VolumeTolerance)
            {
                return null;
            }

            // Set centroid coordinates to their final value
            cv.Centroid /= 4 * cv.Volume;

            // And, just in case you want to know the total volume of the model:
            cv.Volume /= 6;

            return cv;
        }

        // Calculate centroid for all non-empty solids
        // found for the given element. Family instances
        // may have their own non-empty solids, in which
        // case those are used, otherwise the symbol geometry.
        // The symbol geometry could keep track of the
        // instance transform to map it to the actual
        // project location. Instead, we ask for
        // transformed geometry to be returned, so the
        // resulting solids are already in place.
        public static CentroidVolume GetCentroid(Element e, Options opt)
        {
            CentroidVolume cv = null;

            GeometryElement geo = e.get_Geometry(opt);

            Solid s;

            if (null != geo)
            {
                // List of pairs of centroid, volume for each solid

                List<CentroidVolume> a = [];

                if (e is FamilyInstance)
                {
                    geo = geo.GetTransformed(Transform.Identity);
                }

                GeometryInstance inst = null;

                CentroidVolume cv1;

                foreach (GeometryObject obj in geo)
                {
                    s = obj as Solid;

                    if (null != s
                      && 0 < s.Faces.Size
                      && SolidUtils.IsValidForTessellation(s)
                      && (null != (cv1 = GetCentroid(s))))
                    {
                        a.Add(cv1);
                    }
                    inst = obj as GeometryInstance;
                }

                if (0 == a.Count && null != inst)
                {
                    geo = inst.GetSymbolGeometry();

                    foreach (GeometryObject obj in geo)
                    {
                        s = obj as Solid;

                        if (null != s
                          && 0 < s.Faces.Size
                          && SolidUtils.IsValidForTessellation(s)
                          && (null != (cv1 = GetCentroid(s))))
                        {
                            a.Add(cv1);
                        }
                    }
                }

                // Get the total centroid from the partial contributions.
                // Each contribution is weighted with its associated volume,
                // which is factored out again at the end:
                //     C = sum(V_i * C_i) / sum(V_i)

                if (0 < a.Count)
                {
                    cv = new CentroidVolume();

                    foreach (CentroidVolume cv2 in a)
                    {
                        cv.Centroid += cv2.Volume * cv2.Centroid;
                        cv.Volume += cv2.Volume;
                    }

                    if (Math.Abs(cv.Volume) > VolumeTolerance)
                    {
                        cv.Centroid /= cv.Volume;
                    }
                    else
                    {
                        cv = null;
                    }
                }
            }

            return cv;
        }

        public static CentroidVolume GetCentroid(List<Element> elements, Options opt)
        {
            return GetCentroid(elements, opt, null);
        }

        /// <summary>
        /// Combined centre of gravity for a set of elements. Always computes the
        /// volume-weighted result; when <paramref name="density"/> is supplied it also
        /// computes the mass-weighted result (each element weighted by volume x density).
        /// </summary>
        public static CentroidVolume GetCentroid(List<Element> elements, Options opt, IDensityProvider density)
        {
            CentroidVolume cv = new();

            foreach (var element in elements)
            {
                CentroidVolume cv1 = GetCentroid(element, opt);

                if (cv1 == null || !cv1.IsValid)
                {
                    cv.SkippedElementIds.Add(element.Id);
                    continue;
                }

                double mass = 0.0;
                bool massKnown = false;

                if (density != null)
                {
                    mass = ComputeMass(element, cv1.Volume, density, out massKnown);
                }

                cv.Contributions.Add(new ElementContribution
                {
                    Id = element.Id,
                    Volume = cv1.Volume,
                    Centroid = cv1.Centroid,
                    Mass = mass,
                    MassKnown = massKnown,
                    Density = (massKnown && cv1.Volume > VolumeTolerance) ? mass / cv1.Volume : 0.0
                });

                cv.Centroid += cv1.Volume * cv1.Centroid;
                cv.Volume += cv1.Volume;

                if (density != null)
                {
                    if (massKnown)
                    {
                        cv.MassCentroid += mass * cv1.Centroid;
                        cv.Mass += mass;
                    }
                    else
                    {
                        cv.NoDensityElementIds.Add(element.Id);
                    }
                }
            }

            if (Math.Abs(cv.Volume) > VolumeTolerance)
            {
                cv.Centroid /= cv.Volume;
            }
            else
            {
                // Nothing in the selection had usable solid geometry.
                cv.Centroid = XYZ.Zero;
                cv.Volume = 0.0;
            }

            if (Math.Abs(cv.Mass) > MassTolerance)
            {
                cv.MassCentroid /= cv.Mass;
                cv.MassComplete = cv.Contributions.Count > 0 && cv.NoDensityElementIds.Count == 0;
            }
            else
            {
                cv.MassCentroid = XYZ.Zero;
                cv.Mass = 0.0;
                cv.MassComplete = false;
            }

            return cv;
        }

        /// <summary>
        /// Element mass in internal units. Prefers per-material volumes x per-material
        /// density; falls back to (element volume x element density).
        /// </summary>
        private static double ComputeMass(Element e, double elementVolume, IDensityProvider density, out bool known)
        {
            known = false;
            double mass = 0.0;
            bool any = false;

            try
            {
                foreach (ElementId matId in e.GetMaterialIds(false))
                {
                    double materialVolume;
                    try
                    {
                        materialVolume = e.GetMaterialVolume(matId);
                    }
                    catch (Exception)
                    {
                        materialVolume = 0.0;
                    }

                    if (materialVolume <= 0.0)
                    {
                        continue;
                    }

                    double? d = density.GetMaterialDensity(matId);

                    if (d.HasValue && d.Value > 0.0)
                    {
                        mass += materialVolume * d.Value;
                        any = true;
                    }
                }
            }
            catch (Exception)
            {
                // element does not support material queries
            }

            if (!any && elementVolume > VolumeTolerance)
            {
                double? d = density.GetElementDensity(e);

                if (d.HasValue && d.Value > 0.0)
                {
                    mass = elementVolume * d.Value;
                    any = true;
                }
            }

            known = any && mass > 0.0;
            return known ? mass : 0.0;
        }

        public static bool IsPhysicalElement(this Element e)
        {
            if (e.Category == null || e.ViewSpecific)
            {
                return false;
            }

            return e.Category.CategoryType == CategoryType.Model && e.Category.CanAddSubcategory;
        }
    }
}
