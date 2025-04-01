using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace BBI.JD
{

    public class CentroidVolume
    {
        public CentroidVolume()
        {
            Centroid = XYZ.Zero;
            Volume = 0.0;
        }

        public XYZ Centroid { get; set; }
        public double Volume { get; set; }
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

            // Set centroid coordinates to their final value

            cv.Centroid /= 4 * cv.Volume;

            // XYZ diffCentroid = cv.Centroid - solid.ComputeCentroid();

            // And, just in case you want to know 
            // the total volume of the model:

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

                // Get the total centroid from the partial
                // contributions. Each contribution is weighted
                // with its associated volume, which needs to 
                // be factored out again at the end.

                if (0 < a.Count)
                {
                    cv = new CentroidVolume();

                    foreach (CentroidVolume cv2 in a)
                    {
                        cv.Centroid += cv2.Volume * cv2.Centroid;
                        cv.Volume += cv2.Volume;
                    }

                    cv.Centroid /= a.Count * cv.Volume;
                }
            }

            return cv;
        }

        public static CentroidVolume GetCentroid(List<Element> elements, Options opt)
        {
            CentroidVolume cv = new();

            foreach (var element in elements)
            {
                CentroidVolume cv1 = GetCentroid(element, opt);

                if (cv1 != null)
                {
                    cv.Centroid = cv.Centroid.Add(new XYZ(cv1.Centroid.X * cv1.Volume, cv1.Centroid.Y * cv1.Volume, cv1.Centroid.Z * cv1.Volume));
                    cv.Volume += cv1.Volume;
                }
            }

            cv.Centroid = cv.Centroid.Divide(cv.Volume);

            return cv;
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