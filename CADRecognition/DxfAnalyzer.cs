using netDxf;
using netDxf.Entities;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CADRecognition
{
    public static class DxfAnalyzer
    {
        private const int SignatureSamples = 72;

        public static MoldProfile ExtractMold(int moldId, string path)
        {
            var doc = CadDocumentLoader.Load(path);
            var outer = DetectOuterRectangle(doc);
            var features = ExtractHoleFeatures(doc).ToList();
            var feature = features.OrderByDescending(f => f.Area).FirstOrDefault()
                ?? new HoleFeature("Unknown", (outer.MinX, outer.MinY), 1, 1, 1, 1, 0, CreateCircleSignature(1, SignatureSamples));

            var outline = ExtractMoldOutline(doc, outer);
            return new MoldProfile(moldId, path, feature, outline, features);
        }

        public static ProjectProfile ExtractProject(DxfDocument doc)
        {
            var outer = DetectOuterRectangle(doc);
            var holes = ExtractHoleFeatures(doc).ToList();
            var edgeCandidates = ExtractEdgeCandidates(doc, outer);
            var cornerCandidates = ExtractCornerCandidates(doc, outer);
            var contourPaths = ExtractContourPaths(doc, outer);
            return new ProjectProfile(outer, holes, cornerCandidates, edgeCandidates, [], contourPaths);
        }

        public static IReadOnlyList<(double X, double Y)> ExtractOuterContourForDebug(DxfDocument doc)
        {
            var outer = DetectOuterRectangle(doc);
            return new List<(double X, double Y)>
            {
                (outer.MinX, outer.MinY),
                (outer.MinX, outer.MaxY),
                (outer.MaxX, outer.MaxY),
                (outer.MaxX, outer.MinY),
                (outer.MinX, outer.MinY)
            };
        }

        public static RawBounds GetRawBounds(DxfDocument doc)
        {
            var pts = CollectGeometryPoints(doc);
            if (pts.Count == 0)
            {
                return new RawBounds(0, 0, 100, 100);
            }

            return new RawBounds(pts.Min(p => p.X), pts.Min(p => p.Y), pts.Max(p => p.X), pts.Max(p => p.Y));
        }

        public static RectBounds DetectOuterRectangle(DxfDocument doc)
        {
            var raw = GetRawBounds(doc);
            return new RectBounds(raw.MinX, raw.MinY, raw.MaxX, raw.MaxY);
        }

        private static IEnumerable<HoleFeature> ExtractHoleFeatures(DxfDocument doc)
        {
            foreach (var c in doc.Entities.Circles)
            {
                yield return new HoleFeature("Circle", (c.Center.X, c.Center.Y), c.Radius * 2, c.Radius * 2, Math.PI * c.Radius * c.Radius, 2 * Math.PI * c.Radius, 0, CreateCircleSignature(c.Radius, SignatureSamples));
            }

            foreach (var a in doc.Entities.Arcs)
            {
                var sweep = NormalizeArcSweep(a.StartAngle, a.EndAngle);
                if (sweep >= 350.0)
                {
                    yield return new HoleFeature("ArcCircle", (a.Center.X, a.Center.Y), a.Radius * 2, a.Radius * 2, Math.PI * a.Radius * a.Radius, 2 * Math.PI * a.Radius, 0, CreateCircleSignature(a.Radius, SignatureSamples));
                }
            }

            foreach (var pl in doc.Entities.Polylines2D.Where(p => p.Vertexes.Count >= 3))
            {
                var pts = ExpandPolyline2D(pl, 24);
                if (pts.Count < 3)
                {
                    continue;
                }
                var area = Math.Abs(PolygonArea(pts));
                if (area < 1e-6)
                {
                    continue;
                }
                yield return new HoleFeature("Polyline", (pts.Average(p => p.X), pts.Average(p => p.Y)), pts.Max(p => p.X) - pts.Min(p => p.X), pts.Max(p => p.Y) - pts.Min(p => p.Y), area, PolylineLength(pts), 0, CreatePolylineSignature(pts, SignatureSamples));
            }
        }

        private static IReadOnlyList<EdgeCandidate> ExtractEdgeCandidates(DxfDocument doc, RectBounds outer)
        {
            return [];
        }

        private static IReadOnlyList<HoleFeature> ExtractCornerCandidates(DxfDocument doc, RectBounds outer)
        {
            return [];
        }

        private static IReadOnlyList<CornerStepPath> ExtractContourPaths(DxfDocument doc, RectBounds outer)
        {
            var pts = CollectGeometryPoints(doc);
            if (pts.Count < 2)
            {
                return [];
            }

            return new[] { new CornerStepPath("Contour1", pts) };
        }

        private static List<(double X, double Y)> ExtractMoldOutline(DxfDocument doc, RectBounds outer)
        {
            var pts = CollectGeometryPoints(doc);
            if (pts.Count < 2)
            {
                return [];
            }
            var cx = (outer.MinX + outer.MaxX) * 0.5;
            var cy = (outer.MinY + outer.MaxY) * 0.5;
            var ordered = pts.OrderBy(p => Math.Atan2(p.Y - cy, p.X - cx)).ToList();
            if (ordered.Count > 0)
            {
                ordered.Add(ordered[0]);
            }
            return ordered;
        }

        public static List<(double X, double Y)> CollectGeometryPoints(DxfDocument doc)
        {
            var pts = new List<(double X, double Y)>();
            pts.AddRange(doc.Entities.Lines.SelectMany(l => new[] { (l.StartPoint.X, l.StartPoint.Y), (l.EndPoint.X, l.EndPoint.Y) }));
            pts.AddRange(doc.Entities.Circles.Select(c => (c.Center.X, c.Center.Y)));
            pts.AddRange(doc.Entities.Arcs.SelectMany(a => SampleArc(a, 24)));
            pts.AddRange(doc.Entities.Polylines2D.SelectMany(ExpandPolyline2D));
            return pts;
        }

        public static List<(double X, double Y)> ExpandPolyline2D(Polyline2D polyline, int bulgeSamplesPerSegment)
        {
            var pts = new List<(double X, double Y)>();
            foreach (var v in polyline.Vertexes)
            {
                pts.Add((v.Position.X, v.Position.Y));
            }
            return pts;
        }

        public static List<(double X, double Y)> SampleArc(Arc arc, int segments)
        {
            var startDeg = arc.StartAngle;
            var endDeg = arc.EndAngle;
            while (endDeg <= startDeg)
            {
                endDeg += 360.0;
            }
            var pts = new List<(double X, double Y)>(segments + 1);
            for (var i = 0; i <= segments; i++)
            {
                var t = (double)i / segments;
                var deg = startDeg + (endDeg - startDeg) * t;
                var rad = deg * Math.PI / 180.0;
                pts.Add((arc.Center.X + arc.Radius * Math.Cos(rad), arc.Center.Y + arc.Radius * Math.Sin(rad)));
            }
            return pts;
        }

        private static double NormalizeArcSweep(double startAngle, double endAngle)
        {
            var sweep = endAngle - startAngle;
            while (sweep < 0) sweep += 360.0;
            while (sweep > 360.0) sweep -= 360.0;
            return sweep;
        }

        private static double PolygonArea(IReadOnlyList<(double X, double Y)> points)
        {
            double sum = 0;
            for (var i = 0; i < points.Count; i++)
            {
                var j = (i + 1) % points.Count;
                sum += points[i].X * points[j].Y - points[j].X * points[i].Y;
            }
            return sum / 2.0;
        }

        private static double PolylineLength(IReadOnlyList<(double X, double Y)> points)
        {
            double sum = 0;
            for (var i = 1; i < points.Count; i++)
            {
                var dx = points[i].X - points[i - 1].X;
                var dy = points[i].Y - points[i - 1].Y;
                sum += Math.Sqrt(dx * dx + dy * dy);
            }
            return sum;
        }

        private static double[] CreateCircleSignature(double radius, int samples)
        {
            var arr = new double[samples];
            for (var i = 0; i < samples; i++) arr[i] = 1.0;
            return arr;
        }

        private static double[] CreatePolylineSignature(IReadOnlyList<(double X, double Y)> points, int samples)
        {
            var arr = new double[samples];
            for (var i = 0; i < samples; i++) arr[i] = 1.0;
            return arr;
        }
    }
}
