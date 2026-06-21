using System;
using System.IO;

namespace ShapePalette
{
    /// <summary>デバッグ用の簡易ログ。%TEMP%\ShapePalette.log に追記する。</summary>
    internal static class Log
    {
        private static readonly object _lock = new object();
        private static readonly string _path =
            Path.Combine(Path.GetTempPath(), "ShapePalette.log");

        public static void Write(string message)
        {
            try
            {
                lock (_lock)
                {
                    File.AppendAllText(_path,
                        DateTime.Now.ToString("HH:mm:ss.fff") + "  " + message + Environment.NewLine);
                }
            }
            catch { }
        }
    }
}
