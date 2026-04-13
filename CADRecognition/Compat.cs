using System;
using System.Collections.Generic;

namespace CADRecognition
{
    internal static class Compat
    {
        public static double Clamp(double value, double min, double max)
        {
            if (value < min) return min;
            if (value > max) return max;
            return value;
        }

        public static IEnumerable<TSource> DistinctBy<TSource, TKey>(
            this IEnumerable<TSource> source,
            Func<TSource, TKey> keySelector)
        {
            var seen = new HashSet<TKey>();
            foreach (var item in source)
            {
                if (seen.Add(keySelector(item)))
                {
                    yield return item;
                }
            }
        }

        public static bool ContainsIgnoreCase(this string source, string value)
        {
            return source != null && value != null && source.IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
