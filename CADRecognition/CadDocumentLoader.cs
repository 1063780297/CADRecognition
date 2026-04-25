using CADImport;
using netDxf;
using netDxf.Entities;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace CADRecognition
{
    internal static class CadDocumentLoader
    {
        public static DxfDocument Load(string path)
        {
            var ext = System.IO.Path.GetExtension(path);
            if (string.Equals(ext, ".dwg", StringComparison.OrdinalIgnoreCase))
            {
                return LoadDwg(path);
            }

            return DxfDocument.Load(path);
        }

        private static DxfDocument LoadDwg(string path)
        {
            using var editor = new CADImport.CADImportControls.CADEditorControl();
            editor.LoadFile(path);

            var image = editor.Image;
            var entities = image?.CurrentLayout?.Entities?.Cast<object>().ToList() ?? [];
            if (entities.Count == 0)
            {
                throw new InvalidOperationException("DWG 读取成功但未解析出实体，请检查图纸版本或CADImport运行库。");
            }

            var doc = new DxfDocument();
            var stats = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            foreach (var e in entities)
            {
                if (!TryConvertCadImportEntity(e, doc))
                {
                    var key = e.GetType().Name;
                    stats.TryGetValue(key, out var count);
                    stats[key] = count + 1;
                }
            }

            if (doc.Entities == null)
            {
                var details = stats.Count == 0 ? "未识别到可转换实体" : string.Join(", ", stats.Select(kv => $"{kv.Key}:{kv.Value}"));
                throw new InvalidOperationException($"DWG 读取成功但未提取出可绘制实体，请检查图纸内容。{details}");
            }

            return doc;
        }

        private static bool TryConvertCadImportEntity(object entity, DxfDocument doc)
        {
            switch (entity)
            {
                case CADLine ln:
                    doc.Entities.Add(new Line(new Vector3(ln.Point.X, ln.Point.Y, 0), new Vector3(ln.Point1.X, ln.Point1.Y, 0)));
                    return true;
                case CADArc ca:
                    var center = new Vector3(ca.Point.X, ca.Point.Y, 0);
                    var radius = ca.Radius;
                    var start = ReadCadAngleDegrees(ca, "StartAngle", "StartParam");
                    var end = ReadCadAngleDegrees(ca, "EndAngle", "EndParam");
                    if (Math.Abs(end - start) < 1e-6)
                    {
                        end = start + 360.0;
                    }
                    doc.Entities.Add(new Arc(center, radius, start, end));
                    return true;
                case CADCircle cc:
                    doc.Entities.Add(new Circle(new Vector3(cc.Point.X, cc.Point.Y, 0), cc.Radius));
                    return true;
                case object poly when IsPolylineLike(poly):
                    doc.Entities.Add(ToPolyline2D(poly));
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsPolylineLike(object poly)
        {
            var name = poly.GetType().Name;
            return name.Contains("Polyline", StringComparison.OrdinalIgnoreCase);
        }

        private static Polyline2D ToPolyline2D(object poly)
        {
            var points = GetCadPolyVertices(poly);
            var result = new Polyline2D();
            foreach (var pt in points)
            {
                result.Vertexes.Add(new Polyline2DVertex(new Vector2(pt.X, pt.Y)));
            }
            result.IsClosed = IsCadPolylineClosed(poly, points);
            return result;
        }

        private static List<(double X, double Y)> GetCadPolyVertices(object poly)
        {
            var result = new List<(double X, double Y)>();
            var type = poly.GetType();
            var vertexProp = type.GetProperty("Vertexes") ?? type.GetProperty("Vertices") ?? type.GetProperty("Points");
            if (vertexProp?.GetValue(poly) is System.Collections.IEnumerable items)
            {
                foreach (var item in items)
                {
                    if (TryReadPoint2D(item, out var p))
                    {
                        result.Add(p);
                    }
                }
            }
            return result;
        }

        private static bool TryReadPoint2D(object? obj, out (double X, double Y) point)
        {
            point = (0, 0);
            if (obj is null)
            {
                return false;
            }

            var type = obj.GetType();
            var xProp = type.GetProperty("X") ?? type.GetProperty("x") ?? type.GetProperty("PointX");
            var yProp = type.GetProperty("Y") ?? type.GetProperty("y") ?? type.GetProperty("PointY");
            if (xProp is null || yProp is null)
            {
                return false;
            }

            if (!TryConvertToDouble(xProp.GetValue(obj), out var x) || !TryConvertToDouble(yProp.GetValue(obj), out var y))
            {
                return false;
            }

            point = (x, y);
            return true;
        }

        private static bool TryConvertToDouble(object? value, out double result)
        {
            switch (value)
            {
                case double d:
                    result = d;
                    return true;
                case float f:
                    result = f;
                    return true;
                case int i:
                    result = i;
                    return true;
                case long l:
                    result = l;
                    return true;
                case short s:
                    result = s;
                    return true;
                default:
                    if (value is IConvertible convertible)
                    {
                        result = convertible.ToDouble(CultureInfo.InvariantCulture);
                        return true;
                    }
                    result = 0;
                    return false;
            }
        }

        private static bool IsCadPolylineClosed(object poly, IReadOnlyList<(double X, double Y)> points)
        {
            var type = poly.GetType();
            var closedProp = type.GetProperty("IsClosed") ?? type.GetProperty("Closed");
            if (closedProp?.GetValue(poly) is bool closed)
            {
                return closed;
            }

            if (points.Count < 3)
            {
                return false;
            }

            var first = points[0];
            var last = points[^1];
            return Math.Sqrt((first.X - last.X) * (first.X - last.X) + (first.Y - last.Y) * (first.Y - last.Y)) <= 1e-3;
        }

        private static double ReadCadAngleDegrees(object arc, params string[] propertyNames)
        {
            var type = arc.GetType();
            foreach (var name in propertyNames)
            {
                var prop = type.GetProperty(name);
                if (prop is null)
                {
                    continue;
                }

                var value = prop.GetValue(arc);
                if (TryConvertToDouble(value, out var raw))
                {
                    return Math.Abs(raw) <= 2.0 * Math.PI + 1e-6 ? NormalizeDeg(raw * 180.0 / Math.PI) : NormalizeDeg(raw);
                }
            }

            return 0.0;
        }

        private static double NormalizeDeg(double degree)
        {
            var d = degree % 360.0;
            if (d < 0) d += 360.0;
            return d;
        }
    }
}
