using System;
using System.Collections.Generic;
using System.Linq;

namespace CADRecognition
{
    public sealed class MoldMatcher
    {
        public MatchResult Match(ProjectProfile project, IReadOnlyList<MoldProfile> molds)
        {
            var rows = new List<HoleAssignment>();
            var guidePaths = new List<CornerStepPath>();
            if (molds.Count == 0 || project.Holes.Count == 0)
            {
                return new MatchResult(rows, guidePaths);
            }

            var mold1 = molds.FirstOrDefault(m => m.MoldId == 1) ?? molds[0];
            foreach (var hole in project.Holes)
            {
                var best = molds.OrderBy(m => Score(hole, m.Feature)).First();
                rows.Add(new HoleAssignment(
                    hole,
                    best.MoldId,
                    "单次冲压",
                    false,
                    false,
                    $"M{best.MoldId:D2}"));
            }

            return new MatchResult(rows, guidePaths);
        }

        private static double Score(HoleFeature hole, HoleFeature mold)
        {
            var area = Math.Abs(hole.Area - mold.Area) / Math.Max(mold.Area, 1e-6);
            var peri = Math.Abs(hole.Perimeter - mold.Perimeter) / Math.Max(mold.Perimeter, 1e-6);
            return area + peri;
        }
    }
}
