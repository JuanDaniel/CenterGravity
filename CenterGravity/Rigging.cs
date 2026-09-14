using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;

namespace BBI.JD
{
    public class RiggingPointResult
    {
        public int Index { get; set; }
        public XYZ Point { get; set; }

        /// <summary>Vertical load carried by this leg, in Revit internal mass units.</summary>
        public double VerticalLoad { get; set; }

        /// <summary>Share of the total weight, 0..1 (can be negative when a leg goes into uplift).</summary>
        public double Fraction { get; set; }

        /// <summary>Angle of the sling from vertical, radians. 0 when no hook height was given.</summary>
        public double SlingAngle { get; set; }

        /// <summary>Tension along the sling, internal mass units. Equals VerticalLoad when no hook height.</summary>
        public double SlingTension { get; set; }
    }

    public class RiggingResult
    {
        public bool Valid { get; set; }
        public string Message { get; set; }

        public List<RiggingPointResult> Points { get; set; } = new();

        /// <summary>Total lifted weight, internal mass units.</summary>
        public double TotalWeight { get; set; }

        /// <summary>True when the centre of gravity projects inside the support polygon / span.</summary>
        public bool CogInsideSupport { get; set; }

        /// <summary>True when at least one leg computes as negative (uplift).</summary>
        public bool AnyUplift { get; set; }

        /// <summary>Tilt the load would take if rigged symmetrically over the pick points, radians. 0 if no hook height.</summary>
        public double TiltIfSymmetric { get; set; }

        /// <summary>Horizontal distance between the CoG and the centroid of the pick points, internal length.</summary>
        public double CogEccentricity { get; set; }
    }

    /// <summary>
    /// Preliminary rigging checks - NOT a certified lift plan. 2 points use an exact
    /// moment split along the line; 3+ points use the "plane stays plane" distribution
    /// (same idea as a bolt group), which assumes equal-stiffness legs.
    /// </summary>
    public static class RiggingCalculator
    {
        public static RiggingResult Compute(XYZ cog, double totalWeight, IList<XYZ> points, double hookHeight)
        {
            RiggingResult result = new();

            if (points == null || points.Count < 2)
            {
                result.Message = "Pick at least two lift points.";
                return result;
            }

            if (totalWeight <= 0 || cog == null)
            {
                result.Message = "No weight available - weight by mass or set a default density first.";
                return result;
            }

            result.TotalWeight = totalWeight;

            XYZ centroid = Average(points);
            result.CogEccentricity = Distance2D(cog, centroid);

            double[] fractions = points.Count == 2
                ? SplitTwoPoints(cog, points, out bool inside2)
                : SplitPlane(cog, points, centroid, out bool insideN);

            bool inside = points.Count == 2
                ? IsBetween(cog, points[0], points[1])
                : PointInPolygon(cog, OrderAroundCentroid(points, centroid));

            result.CogInsideSupport = inside;

            for (int i = 0; i < points.Count; i++)
            {
                double vertical = fractions[i] * totalWeight;

                RiggingPointResult p = new()
                {
                    Index = i + 1,
                    Point = points[i],
                    Fraction = fractions[i],
                    VerticalLoad = vertical,
                    SlingAngle = 0.0,
                    SlingTension = vertical
                };

                if (hookHeight > 1e-6)
                {
                    XYZ hook = new(cog.X, cog.Y, cog.Z + hookHeight);
                    XYZ leg = points[i] - hook;
                    double legLength = leg.GetLength();
                    double verticalDrop = Math.Abs(hook.Z - points[i].Z);

                    if (legLength > 1e-6 && verticalDrop > 1e-6)
                    {
                        double cos = verticalDrop / legLength;
                        p.SlingAngle = Math.Acos(Math.Min(1.0, Math.Max(-1.0, cos)));
                        p.SlingTension = Math.Abs(vertical) / Math.Max(cos, 1e-6);
                    }
                }

                result.Points.Add(p);
                if (vertical < 0) result.AnyUplift = true;
            }

            if (hookHeight > 1e-6)
            {
                result.TiltIfSymmetric = Math.Atan2(result.CogEccentricity, hookHeight);
            }

            result.Valid = true;
            return result;
        }

        private static double[] SplitTwoPoints(XYZ cog, IList<XYZ> pts, out bool inside)
        {
            XYZ a = pts[0];
            XYZ b = pts[1];
            XYZ ab = new(b.X - a.X, b.Y - a.Y, 0);
            double len2 = ab.X * ab.X + ab.Y * ab.Y;

            double t = len2 > 1e-9
                ? ((cog.X - a.X) * ab.X + (cog.Y - a.Y) * ab.Y) / len2
                : 0.5;

            inside = t >= 0 && t <= 1;
            return new[] { 1.0 - t, t };
        }

        private static double[] SplitPlane(XYZ cog, IList<XYZ> pts, XYZ centroid, out bool inside)
        {
            int n = pts.Count;
            double ix = pts.Sum(p => (p.X - centroid.X) * (p.X - centroid.X));
            double iy = pts.Sum(p => (p.Y - centroid.Y) * (p.Y - centroid.Y));

            double ex = cog.X - centroid.X;
            double ey = cog.Y - centroid.Y;

            double[] f = new double[n];
            for (int i = 0; i < n; i++)
            {
                double term = 1.0 / n;
                if (ix > 1e-9) term += ex * (pts[i].X - centroid.X) / ix;
                if (iy > 1e-9) term += ey * (pts[i].Y - centroid.Y) / iy;
                f[i] = term;
            }

            // renormalise so the shares always sum to 1 despite rounding
            double sum = f.Sum();
            if (Math.Abs(sum) > 1e-9)
            {
                for (int i = 0; i < n; i++) f[i] /= sum;
            }

            inside = true;
            return f;
        }

        private static XYZ Average(IList<XYZ> pts)
        {
            double x = 0, y = 0, z = 0;
            foreach (XYZ p in pts) { x += p.X; y += p.Y; z += p.Z; }
            return new XYZ(x / pts.Count, y / pts.Count, z / pts.Count);
        }

        private static double Distance2D(XYZ a, XYZ b)
        {
            double dx = a.X - b.X;
            double dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        private static bool IsBetween(XYZ cog, XYZ a, XYZ b)
        {
            XYZ ab = new(b.X - a.X, b.Y - a.Y, 0);
            double len2 = ab.X * ab.X + ab.Y * ab.Y;
            if (len2 < 1e-9) return false;
            double t = ((cog.X - a.X) * ab.X + (cog.Y - a.Y) * ab.Y) / len2;
            return t >= 0 && t <= 1;
        }

        private static List<XYZ> OrderAroundCentroid(IList<XYZ> pts, XYZ centroid)
        {
            return pts
                .OrderBy(p => Math.Atan2(p.Y - centroid.Y, p.X - centroid.X))
                .ToList();
        }

        private static bool PointInPolygon(XYZ p, List<XYZ> poly)
        {
            bool inside = false;
            int n = poly.Count;

            for (int i = 0, j = n - 1; i < n; j = i++)
            {
                bool crosses = (poly[i].Y > p.Y) != (poly[j].Y > p.Y);
                if (crosses)
                {
                    double x = (poly[j].X - poly[i].X) * (p.Y - poly[i].Y) / (poly[j].Y - poly[i].Y) + poly[i].X;
                    if (p.X < x) inside = !inside;
                }
            }

            return inside;
        }
    }
}
