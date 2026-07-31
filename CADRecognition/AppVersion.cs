using System.Reflection;

namespace CADRecognition
{
    /// <summary>
    /// 应用程序版本信息管理
    /// </summary>
    public static class AppVersion
    {
        /// <summary>
        /// 当前版本号
        /// </summary>
        public static string Version => "2.0.10";

        /// <summary>
        /// 版本详细信息，包含构建日期等
        /// </summary>
        public static string FullVersion => $"v{Version}";

        /// <summary>
        /// 获取带构建信息的完整版本字符串
        /// </summary>
        public static string GetFullInfo()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var buildDate = System.IO.File.GetLastWriteTime(assembly.Location);
            return $"{FullVersion} (Build: {buildDate:yyyy-MM-dd HH:mm:ss})";
        }

        /// <summary>
        /// 获取版本号的主版本
        /// </summary>
        public static int Major => 2;

        /// <summary>
        /// 获取版本号的次版本
        /// </summary>
        public static int Minor => 0;

        /// <summary>
        /// 获取版本号的修订版本
        /// </summary>
        public static int Patch => 1;
    }
}
