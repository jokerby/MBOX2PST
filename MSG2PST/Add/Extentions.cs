using System.IO;
using System.Linq;

namespace MSG2PST.Add
{
    public static class Extentions
    {
        public static string Tail(this string source, int length) => length >= source.Length ? source : source.Substring(source.Length - 4);

        public static string TrimIllegalFromPath(this string source) => (new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars())).Aggregate(source,
                               (current, c) => current.Replace(c.ToString(), ""));
    }
}