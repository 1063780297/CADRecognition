using System.Collections.Generic;
using System.IO;
using System.Text.Json;

namespace CADRecognition
{
    /// <summary>持久化上次会话中的台1/台2模具文件路径，便于下次启动自动加载。</summary>
    internal static class LastMoldSessionSettings
    {
        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        private static string GetFilePath()
        {
            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "CADRecognition");
            return Path.Combine(dir, "last-molds.json");
        }

        internal sealed class Dto
        {
            public List<string> Stage1 { get; set; } = [];
            public List<string> Stage2 { get; set; } = [];
        }

        public static Dto Load()
        {
            var path = GetFilePath();
            if (!File.Exists(path))
            {
                return new Dto();
            }

            try
            {
                var json = File.ReadAllText(path);
                var dto = JsonSerializer.Deserialize<Dto>(json, JsonOptions);
                return dto ?? new Dto();
            }
            catch
            {
                return new Dto();
            }
        }

        public static void Save(IReadOnlyList<string> stage1, IReadOnlyList<string> stage2)
        {
            try
            {
                var path = GetFilePath();
                var dir = Path.GetDirectoryName(path);
                if (!string.IsNullOrEmpty(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                var dto = new Dto
                {
                    Stage1 = [.. stage1],
                    Stage2 = [.. stage2]
                };
                File.WriteAllText(path, JsonSerializer.Serialize(dto, JsonOptions));
            }
            catch
            {
                // 忽略持久化失败，避免影响关闭应用
            }
        }
    }
}
