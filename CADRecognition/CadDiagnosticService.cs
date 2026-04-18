using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using netDxf;
using netDxf.Entities;

namespace CADRecognition
{
    internal static class CadDiagnosticService
    {
        internal sealed record CadSummary(
            string FilePath,
            string Format,
            int LineCount,
            int CircleCount,
            int ArcCount,
            int Polyline2DCount,
            int InsertCount,
            int EntityCount,
            RectBounds Bounds)
        {
            public string ShortText =>
                $"{Path.GetFileName(FilePath)} | {Format} | 实体:{EntityCount} 线:{LineCount} 圆:{CircleCount} 弧:{ArcCount} 多段线:{Polyline2DCount} 块:{InsertCount} | 外接框:{Bounds.MinX:F2},{Bounds.MinY:F2}~{Bounds.MaxX:F2},{Bounds.MaxY:F2}";
        }

        internal sealed record DiagnosticReport(string ShortSummary, string FullText);

        public static DiagnosticReport BuildReport(string currentPath, DxfDocument currentDoc, IEnumerable<string> comparePaths, Func<string, DxfDocument> loader)
        {
            var lines = new List<string>();
            var current = Summarize(currentPath, currentDoc);
            lines.Add("当前工程");
            lines.Add(current.ShortText);
            lines.Add(string.Empty);

            foreach (var path in comparePaths ?? Enumerable.Empty<string>())
            {
                if (string.Equals(path, currentPath, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (!File.Exists(path))
                {
                    lines.Add($"{Path.GetFileName(path)} | 文件不存在");
                    continue;
                }

                try
                {
                    var doc = loader(path);
                    var summary = Summarize(path, doc);
                    lines.Add("对照图纸");
                    lines.Add(summary.ShortText);
                    lines.Add(GetDiffLine(current, summary));
                    lines.Add(string.Empty);
                }
                catch (Exception ex)
                {
                    lines.Add($"{Path.GetFileName(path)} | 读取失败：{ex.Message}");
                    lines.Add(string.Empty);
                }
            }

            if (lines.Count == 2)
            {
                lines.Add("没有找到可对照的模具图纸。");
            }

            var summaryText = string.Join(Environment.NewLine, lines.Where(x => !string.IsNullOrWhiteSpace(x)).Take(4));
            return new DiagnosticReport(summaryText, string.Join(Environment.NewLine, lines));
        }

        private static CadSummary Summarize(string path, DxfDocument doc)
        {
            var bounds = GetBounds(doc);
            var lineCount = doc.Entities.Lines.Count();
            var circleCount = doc.Entities.Circles.Count();
            var arcCount = doc.Entities.Arcs.Count();
            var polyCount = doc.Entities.Polylines2D.Count();
            var insertCount = doc.Entities.Inserts.Count();
            var entityCount = lineCount + circleCount + arcCount + polyCount + insertCount;

            return new CadSummary(
                path,
                Path.GetExtension(path).TrimStart('.').ToUpperInvariant(),
                lineCount,
                circleCount,
                arcCount,
                polyCount,
                insertCount,
                entityCount,
                bounds);
        }

        private static string GetDiffLine(CadSummary a, CadSummary b)
        {
            var entityDiff = b.EntityCount - a.EntityCount;
            var lineDiff = b.LineCount - a.LineCount;
            var circleDiff = b.CircleCount - a.CircleCount;
            var arcDiff = b.ArcCount - a.ArcCount;
            var polyDiff = b.Polyline2DCount - a.Polyline2DCount;
            var insertDiff = b.InsertCount - a.InsertCount;
            var dx = (b.Bounds.MinX + b.Bounds.MaxX) / 2.0 - (a.Bounds.MinX + a.Bounds.MaxX) / 2.0;
            var dy = (b.Bounds.MinY + b.Bounds.MaxY) / 2.0 - (a.Bounds.MinY + a.Bounds.MaxY) / 2.0;
            return string.Format(CultureInfo.InvariantCulture,
                "差异 | 实体:{0:+0;-0;0} 线:{1:+0;-0;0} 圆:{2:+0;-0;0} 弧:{3:+0;-0;0} 多段线:{4:+0;-0;0} 块:{5:+0;-0;0} | 中心偏移:{6:F2},{7:F2}",
                entityDiff, lineDiff, circleDiff, arcDiff, polyDiff, insertDiff, dx, dy);
        }

        private static RectBounds GetBounds(DxfDocument doc)
        {
            var points = new List<(double X, double Y)>();
            foreach (var l in doc.Entities.Lines)
            {
                points.Add((l.StartPoint.X, l.StartPoint.Y));
                points.Add((l.EndPoint.X, l.EndPoint.Y));
            }
            foreach (var c in doc.Entities.Circles)
            {
                points.Add((c.Center.X - c.Radius, c.Center.Y - c.Radius));
                points.Add((c.Center.X + c.Radius, c.Center.Y + c.Radius));
            }
            foreach (var a in doc.Entities.Arcs)
            {
                points.Add((a.Center.X - a.Radius, a.Center.Y - a.Radius));
                points.Add((a.Center.X + a.Radius, a.Center.Y + a.Radius));
            }

            if (points.Count == 0)
            {
                return new RectBounds(0, 0, 0, 0);
            }

            return new RectBounds(
                points.Min(p => p.X),
                points.Min(p => p.Y),
                points.Max(p => p.X),
                points.Max(p => p.Y));
        }
    }
}
