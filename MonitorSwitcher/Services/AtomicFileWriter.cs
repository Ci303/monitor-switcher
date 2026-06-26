using System;
using System.IO;

namespace WorkMonitorSwitcher.Services
{
    internal static class AtomicFileWriter
    {
        public static void WriteAllText(string path, string contents)
        {
            if (string.IsNullOrWhiteSpace(path))
                throw new ArgumentException("A file path is required.", nameof(path));

            var fullPath = Path.GetFullPath(path);
            var directory = Path.GetDirectoryName(fullPath);
            if (string.IsNullOrWhiteSpace(directory))
                throw new ArgumentException("The file path must include a directory.", nameof(path));

            Directory.CreateDirectory(directory);

            var tempPath = Path.Combine(
                directory,
                $".{Path.GetFileName(fullPath)}.{Guid.NewGuid():N}.tmp");
            var backupPath = fullPath + ".bak";

            try
            {
                File.WriteAllText(tempPath, contents);

                if (File.Exists(fullPath))
                {
                    TryDelete(backupPath);
                    File.Replace(tempPath, fullPath, backupPath, ignoreMetadataErrors: true);
                }
                else
                {
                    File.Move(tempPath, fullPath);
                }
            }
            catch
            {
                TryDelete(tempPath);
                throw;
            }
        }

        private static void TryDelete(string path)
        {
            try
            {
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                    File.Delete(path);
            }
            catch
            {
                // Best effort cleanup only.
            }
        }
    }
}
