using netDxf;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Brushes = System.Windows.Media.Brushes;
using Color = System.Windows.Media.Color;
using Pen = System.Windows.Media.Pen;
using Point = System.Windows.Point;

namespace CADRecognition
{
    public static class CadPreviewRenderService
    {
        public static void DrawToCanvas(Canvas canvas, DxfDocument document, RawBounds bounds, double viewWidth, double viewHeight, double scale, double margin)
        {
            var unifiedStroke = new SolidColorBrush(Color.FromRgb(144, 238, 144));

            foreach (var line in document.Entities.Lines)
            {
                canvas.Children.Add(new Line
                {
                    X1 = (line.StartPoint.X - bounds.MinX) * scale + margin,
                    Y1 = viewHeight - ((line.StartPoint.Y - bounds.MinY) * scale + margin),
                    X2 = (line.EndPoint.X - bounds.MinX) * scale + margin,
                    Y2 = viewHeight - ((line.EndPoint.Y - bounds.MinY) * scale + margin),
                    Stroke = unifiedStroke,
                    StrokeThickness = 1
                });
            }

            foreach (var circle in document.Entities.Circles)
            {
                var r = circle.Radius * scale;
                var x = (circle.Center.X - bounds.MinX) * scale + margin - r;
                var y = viewHeight - ((circle.Center.Y - bounds.MinY) * scale + margin) - r;
                var el = new Ellipse
                {
                    Width = r * 2,
                    Height = r * 2,
                    Stroke = unifiedStroke,
                    StrokeThickness = 1
                };
                Canvas.SetLeft(el, x);
                Canvas.SetTop(el, y);
                canvas.Children.Add(el);
            }

            foreach (var poly in document.Entities.Polylines2D)
            {
                var sampled = DxfAnalyzer.ExpandPolyline2D(poly, 24);
                if (sampled.Count < 2)
                {
                    continue;
                }

                var points = new PointCollection(sampled.Select(v =>
                    new Point(
                        (v.X - bounds.MinX) * scale + margin,
                        viewHeight - ((v.Y - bounds.MinY) * scale + margin))));
                if (poly.IsClosed)
                {
                    canvas.Children.Add(new Polygon
                    {
                        Points = points,
                        Stroke = unifiedStroke,
                        StrokeThickness = 1,
                        Fill = Brushes.Transparent
                    });
                }
                else
                {
                    canvas.Children.Add(new Polyline
                    {
                        Points = points,
                        Stroke = unifiedStroke,
                        StrokeThickness = 1
                    });
                }
            }

            foreach (var arc in document.Entities.Arcs)
            {
                var sampled = DxfAnalyzer.SampleArc(arc, 32);
                if (sampled.Count < 2)
                {
                    continue;
                }

                var points = new PointCollection(sampled.Select(p =>
                    new Point(
                        (p.X - bounds.MinX) * scale + margin,
                        viewHeight - ((p.Y - bounds.MinY) * scale + margin))));
                canvas.Children.Add(new Polyline
                {
                    Points = points,
                    Stroke = unifiedStroke,
                    StrokeThickness = 1
                });
            }
        }

        public static void DrawToDrawingContext(DrawingContext dc, DxfDocument document, RawBounds bounds, double width, double height, double scale, double margin)
        {
            var linePen = new Pen(new SolidColorBrush(Color.FromRgb(80, 210, 120)), 1);
            var circlePen = new Pen(new SolidColorBrush(Color.FromRgb(240, 200, 80)), 1);
            var polyPen = new Pen(new SolidColorBrush(Color.FromRgb(80, 180, 255)), 1);
            var arcPen = new Pen(new SolidColorBrush(Color.FromRgb(255, 140, 0)), 1);

            Point Map(double x, double y) => new(
                (x - bounds.MinX) * scale + margin,
                height - ((y - bounds.MinY) * scale + margin));

            dc.DrawRectangle(new SolidColorBrush(Color.FromRgb(17, 17, 17)), null, new Rect(0, 0, width, height));

            foreach (var l in document.Entities.Lines)
            {
                dc.DrawLine(linePen, Map(l.StartPoint.X, l.StartPoint.Y), Map(l.EndPoint.X, l.EndPoint.Y));
            }

            foreach (var c in document.Entities.Circles)
            {
                var center = Map(c.Center.X, c.Center.Y);
                dc.DrawEllipse(null, circlePen, center, c.Radius * scale, c.Radius * scale);
            }

            foreach (var p in document.Entities.Polylines2D)
            {
                var pts = DxfAnalyzer.ExpandPolyline2D(p, 24);
                if (pts.Count < 2)
                {
                    continue;
                }

                var geo = new StreamGeometry();
                using var g = geo.Open();
                g.BeginFigure(Map(pts[0].X, pts[0].Y), false, p.IsClosed);
                g.PolyLineTo(pts.Skip(1).Select(t => Map(t.X, t.Y)).ToList(), true, false);
                geo.Freeze();
                dc.DrawGeometry(null, polyPen, geo);
            }

            foreach (var a in document.Entities.Arcs)
            {
                var pts = DxfAnalyzer.SampleArc(a, 32);
                if (pts.Count < 2)
                {
                    continue;
                }

                var geo = new StreamGeometry();
                using var g = geo.Open();
                g.BeginFigure(Map(pts[0].X, pts[0].Y), false, false);
                g.PolyLineTo(pts.Skip(1).Select(t => Map(t.X, t.Y)).ToList(), true, false);
                geo.Freeze();
                dc.DrawGeometry(null, arcPen, geo);
            }
        }
    }
}
