using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using IWshRuntimeLibrary;

namespace RecentFoldersBuilder
{
    internal class Program
    {
        private const int MaxFolders = 30;

        // Debounce settings
        private static readonly TimeSpan DebounceDelay = TimeSpan.FromMilliseconds(800);
        private static Timer _debounceTimer;
        private static readonly object _debounceLock = new object();

        private static string _recentDir;
        private static string _outputDir;

        static int Main(string[] args)
        {
            try
            {
                _recentDir = GetRecentDirectoryPath();
                _outputDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                    "Recent Folders"
                );

                Directory.CreateDirectory(_outputDir);
                EnsureCustomFolderIcon(_outputDir);

                bool watchMode = args.Any(a => string.Equals(a, "--watch", StringComparison.OrdinalIgnoreCase));

                // Initial build
                RefreshOnce();

                if (!watchMode)
                    return 0;

                // Watch mode: keep running, refresh on changes
                RunWatcher();

                return 0;
            }
            catch
            {
                return 1;
            }
        }

        private static void RunWatcher()
        {
            using (var watcher = new FileSystemWatcher(_recentDir, "*.lnk"))
            {
                watcher.IncludeSubdirectories = false;
                watcher.NotifyFilter =
                    NotifyFilters.FileName |
                    NotifyFilters.LastWrite |
                    NotifyFilters.CreationTime;

                watcher.Created += (s, e) => ScheduleRefresh();
                watcher.Changed += (s, e) => ScheduleRefresh();
                watcher.Renamed += (s, e) => ScheduleRefresh();
                watcher.Deleted += (s, e) => ScheduleRefresh();

                watcher.EnableRaisingEvents = true;

                // Keep process alive
                using (var quitEvent = new ManualResetEventSlim(false))
                {
                    quitEvent.Wait();
                }
            }
        }

        private static void ScheduleRefresh()
        {
            lock (_debounceLock)
            {
                if (_debounceTimer == null)
                {
                    _debounceTimer = new Timer(_ => SafeRefresh(), null, DebounceDelay, Timeout.InfiniteTimeSpan);
                }
                else
                {
                    _debounceTimer.Change(DebounceDelay, Timeout.InfiniteTimeSpan);
                }
            }
        }

        private static void SafeRefresh()
        {
            try { RefreshOnce(); }
            catch { /* silent */ }
        }

        private static void RefreshOnce()
        {
            var candidates = CollectRecentFolders(_recentDir);

            var top = candidates
                .OrderByDescending(x => x.LastSeenUtc)
                .Take(MaxFolders)
                .ToList();

            RebuildOutputFolder(_outputDir, top);
        }

        private static string GetRecentDirectoryPath()
        {
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            return Path.Combine(appData, "Microsoft", "Windows", "Recent");
        }

        private static List<RecentFolder> CollectRecentFolders(string recentDir)
        {
            var resultByFolder = new Dictionary<string, RecentFolder>(StringComparer.OrdinalIgnoreCase);

            if (!Directory.Exists(recentDir))
                return new List<RecentFolder>();

            var lnkFiles = Directory.EnumerateFiles(recentDir, "*.lnk", SearchOption.TopDirectoryOnly);
            var shell = new WshShell();

            foreach (var lnkPath in lnkFiles)
            {
                try
                {
                    var lnkInfo = new FileInfo(lnkPath);
                    var lastSeenUtc = lnkInfo.LastWriteTimeUtc;

                    var shortcut = (IWshShortcut)shell.CreateShortcut(lnkPath);
                    var target = shortcut.TargetPath;

                    if (string.IsNullOrWhiteSpace(target))
                        continue;

                    string folderPath = null;

                    if (Directory.Exists(target))
                    {
                        folderPath = Path.GetFullPath(target);
                    }
                    else if (System.IO.File.Exists(target))
                    {
                        var parent = Path.GetDirectoryName(target);
                        if (!string.IsNullOrWhiteSpace(parent) && Directory.Exists(parent))
                            folderPath = Path.GetFullPath(parent);
                    }
                    else
                    {
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(folderPath))
                        continue;

                    if (resultByFolder.TryGetValue(folderPath, out var existing))
                    {
                        if (lastSeenUtc > existing.LastSeenUtc)
                            existing.LastSeenUtc = lastSeenUtc;
                    }
                    else
                    {
                        resultByFolder[folderPath] = new RecentFolder
                        {
                            FolderPath = folderPath,
                            LastSeenUtc = lastSeenUtc
                        };
                    }
                }
                catch
                {
                    // ignore
                }
            }

            return resultByFolder.Values.ToList();
        }

        private static void RebuildOutputFolder(string outputDir, List<RecentFolder> topFolders)
        {
            foreach (var file in Directory.EnumerateFiles(outputDir, "*.lnk", SearchOption.TopDirectoryOnly))
            {
                try { System.IO.File.Delete(file); } catch { }
            }

            var shell = new WshShell();

            for (int i = 0; i < topFolders.Count; i++)
            {
                var item = topFolders[i];
                var folderName = GetFriendlyFolderName(item.FolderPath);

                var prefix = (i + 1).ToString("D2");
                var fileName = $"{prefix} - {SanitizeFileName(folderName)}.lnk";
                var shortcutPath = Path.Combine(outputDir, fileName);

                try
                {
                    var shortcut = (IWshShortcut)shell.CreateShortcut(shortcutPath);
                    shortcut.TargetPath = item.FolderPath;
                    shortcut.WorkingDirectory = item.FolderPath;
                    shortcut.Description = $"Recent folder: {item.FolderPath}";
                    shortcut.Save();
                }
                catch { }
            }
        }

        private static string GetFriendlyFolderName(string folderPath)
        {
            try
            {
                var name = new DirectoryInfo(folderPath).Name;
                if (!string.IsNullOrWhiteSpace(name))
                    return name;

                var root = Path.GetPathRoot(folderPath);
                if (!string.IsNullOrWhiteSpace(root))
                    return root.TrimEnd('\\').Replace(":", "");
            }
            catch { }

            return folderPath;
        }

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "Folder";

            foreach (var c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');

            name = name.Trim();
            if (name.Length > 120)
                name = name.Substring(0, 120);

            return string.IsNullOrWhiteSpace(name) ? "Folder" : name;
        }

        private sealed class RecentFolder
        {
            public string FolderPath { get; set; }
            public DateTime LastSeenUtc { get; set; }
        }

        private static void EnsureCustomFolderIcon(string folderPath)
        {
            try
            {
                // 1) Locate icon next to exe (installed by Inno Setup into {app})
                var appDir = AppDomain.CurrentDomain.BaseDirectory;
                var sourceIconPath = Path.Combine(appDir, "rf.ico");

                if (!System.IO.File.Exists(sourceIconPath))
                    return; // icon not present -> skip silently

                // 2) Copy icon into the target folder (use a hidden file name)
                var targetIconPath = Path.Combine(folderPath, ".rf.ico");
                if (!System.IO.File.Exists(targetIconPath))
                {
                    System.IO.File.Copy(sourceIconPath, targetIconPath, overwrite: true);
                }

                // 3) Create desktop.ini referencing the local icon
                var desktopIniPath = Path.Combine(folderPath, "desktop.ini");

                var iniContent =
                    "[.ShellClassInfo]\r\n" +
                    "IconResource=.rf.ico,0\r\n";

                System.IO.File.WriteAllText(desktopIniPath, iniContent);

                // 4) Set required attributes
                // desktop.ini must be Hidden+System
                SetAttributes(desktopIniPath, FileAttributes.Hidden | FileAttributes.System);

                // icon file hidden (optional but nice)
                SetAttributes(targetIconPath, FileAttributes.Hidden);

                // folder must have ReadOnly or System flag for Explorer to honor desktop.ini icon
                var di = new DirectoryInfo(folderPath);
                di.Attributes = di.Attributes | FileAttributes.ReadOnly | FileAttributes.System;
            }
            catch
            {
                // silent utility
            }
        }

        private static void SetAttributes(string path, FileAttributes attrs)
        {
            try
            {
                var current = System.IO.File.GetAttributes(path);
                System.IO.File.SetAttributes(path, current | attrs);
            }
            catch
            {
                // ignore
            }
        }

    }
}