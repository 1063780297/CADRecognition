using System.Windows.Media;

namespace CADRecognition
{
    public sealed record ProjectProfile(
        RectBounds OuterRectangle,
        IReadOnlyList<HoleFeature> Holes,
        IReadOnlyList<HoleFeature> CornerCandidates,
        IReadOnlyList<EdgeCandidate> EdgeCandidates,
        IReadOnlyList<CornerStepPath> CornerStepPaths,
        IReadOnlyList<CornerStepPath> ContourPaths);

    public sealed record MoldProfile(
        int MoldId,
        string FilePath,
        HoleFeature Feature,
        IReadOnlyList<(double X, double Y)> OutlinePoints,
        IReadOnlyList<HoleFeature>? CandidateFeatures = null);
    public sealed record MatchResult(
        IReadOnlyList<HoleAssignment> HoleAssignments,
        IReadOnlyList<CornerStepPath>? GuidePaths = null);

    public sealed record HoleAssignment(
        HoleFeature Hole,
        int MoldId,
        string PositionRelation,
        bool IsCornerCandidate,
        bool IsEdgeHole,
        string TopCandidates,
        string AreaRatioInfo = "",
        string FailureReason = "",
        double RotationDeg = 0,
        bool IsMirrored = false);

    public sealed record HoleFeature(string HoleType, (double X, double Y) Centroid, double Width, double Height, double Area, double Perimeter, double Rotation, double[] Signature);

    public sealed record EdgeCandidate(
        string Side,
        IReadOnlyList<(double X, double Y)> Points,
        (double X, double Y) Centroid,
        double Width,
        double Height,
        double Perimeter,
        double[] Signature);

    public sealed record CornerStepPath(string CornerName, IReadOnlyList<(double X, double Y)> Points);

    public sealed record RawBounds(double MinX, double MinY, double MaxX, double MaxY)
    {
        public double Width => MaxX - MinX;
        public double Height => MaxY - MinY;
    }

    public sealed record RectBounds(double MinX, double MinY, double MaxX, double MaxY)
    {
        public double Width => MaxX - MinX;
        public double Height => MaxY - MinY;
        public double Area => Width * Height;
        public IReadOnlyList<RectCorner> Corners => new[]
        {
            new RectCorner("左下", MinX, MinY),
            new RectCorner("左上", MinX, MaxY),
            new RectCorner("右下", MaxX, MinY),
            new RectCorner("右上", MaxX, MaxY)
        };
    }

    public sealed record RectCorner(string Name, double X, double Y);

    public sealed class MoldRow
    {
        public ImageSource? MoldPreview { get; set; }
        public string MoldCode { get; set; } = string.Empty;
        public string MoldName { get; set; } = string.Empty;
        public int UsedCount { get; set; }
        public string MatchType { get; set; } = string.Empty;
        public string Remark { get; set; } = string.Empty;
    }

    public sealed class PositionRow
    {
        public int Index { get; set; }
        public string HoleType { get; set; } = string.Empty;
        public int MoldId { get; set; }
        public string MoldCode { get; set; } = string.Empty;
        public double PosX { get; set; }
        public double PosY { get; set; }
        public double AbsX { get; set; }
        public double AbsY { get; set; }
        public string PositionRelation { get; set; } = string.Empty;
        public string IsCornerCandidate { get; set; } = string.Empty;
        public string IsEdgeHole { get; set; } = string.Empty;
        public string TopCandidates { get; set; } = string.Empty;
        public string AreaRatio { get; set; } = string.Empty;
        public string FailureReason { get; set; } = string.Empty;
    }
}
