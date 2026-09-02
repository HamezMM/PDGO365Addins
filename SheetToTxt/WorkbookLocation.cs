using System;
using System.IO;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;

namespace SheetToTxt
{
    /// <summary>
    /// Where a workbook lives, resolved to something the exporter can write next to.
    /// A normal file gives a local folder directly; a OneDrive/SharePoint workbook that
    /// is synced on this machine is mapped back to its local folder; a cloud-only
    /// workbook (or one never saved) has no local folder and the caller must ask the
    /// user where to put the export.
    /// </summary>
    internal sealed class WorkbookLocation
    {
        private WorkbookLocation(bool hasLocalFolder, string folder, string baseName, string fullName)
        {
            HasLocalFolder = hasLocalFolder;
            Folder = folder;
            BaseName = baseName;
            FullName = fullName;
        }

        /// <summary>
        /// True when the workbook resolves to a real local directory we can write the
        /// <c>.txt</c> into. False for a cloud-only workbook or one that has never been
        /// saved.
        /// </summary>
        public bool HasLocalFolder { get; }

        /// <summary>Local directory containing the workbook. Empty when <see cref="HasLocalFolder"/> is false.</summary>
        public string Folder { get; }

        /// <summary>Workbook file name without extension, e.g. <c>Budget</c>. Always set.</summary>
        public string BaseName { get; }

        /// <summary>Full path (or URL) of the workbook, or just its name if unsaved.</summary>
        public string FullName { get; }

        /// <summary>Default <c>.txt</c> path next to the workbook. Only meaningful when <see cref="HasLocalFolder"/>.</summary>
        public string TargetTextPath => Path.Combine(Folder, BaseName + ".txt");

        public static WorkbookLocation FromWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            string fileName = workbook.Name ?? "Workbook.xlsx";
            string baseName = Path.GetFileNameWithoutExtension(fileName);
            if (string.IsNullOrEmpty(baseName)) baseName = "Workbook";

            // Path is "" until the first save. A locally-synced OneDrive/SharePoint file
            // reports a normal file-system path; a cloud-only file reports an https URL,
            // which Path.Combine / SaveAs cannot use.
            string path = (workbook.Path ?? string.Empty).Trim();

            string localFolder = null;
            if (path.Length > 0)
            {
                if (Directory.Exists(path))
                    localFolder = path;
                else if (LooksLikeUrl(path))
                    localFolder = TryResolveSyncedFolder(path);
            }

            bool hasLocal = localFolder != null;
            string fullName = hasLocal
                ? (SafeFullName(workbook) ?? Path.Combine(localFolder, fileName))
                : (path.Length > 0 ? path.TrimEnd('/') + "/" + fileName : fileName);

            return new WorkbookLocation(hasLocal, localFolder ?? string.Empty, baseName, fullName);
        }

        private static string SafeFullName(Excel.Workbook workbook)
        {
            try { return workbook.FullName; } catch { return null; }
        }

        private static bool LooksLikeUrl(string s) =>
            s.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
            s.StartsWith("https://", StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// Maps a SharePoint/OneDrive folder URL to its locally-synced path using the
        /// OneDrive sync-engine registry entries
        /// (<c>HKCU\Software\SyncEngines\Providers\OneDrive\*</c>: <c>MountPoint</c> +
        /// <c>UrlNamespace</c>). Returns null when the library is not synced here or the
        /// registry shape is not what we expect.
        /// </summary>
        private static string TryResolveSyncedFolder(string folderUrl)
        {
            try
            {
                string wanted = Normalize(folderUrl);

                using (RegistryKey providers =
                    Registry.CurrentUser.OpenSubKey(@"Software\SyncEngines\Providers\OneDrive"))
                {
                    if (providers == null) return null;

                    foreach (string name in providers.GetSubKeyNames())
                    {
                        using (RegistryKey provider = providers.OpenSubKey(name))
                        {
                            string mount = provider?.GetValue("MountPoint") as string;
                            string urlNamespace = provider?.GetValue("UrlNamespace") as string;
                            if (string.IsNullOrEmpty(mount) || string.IsNullOrEmpty(urlNamespace)) continue;

                            string prefix = Normalize(urlNamespace);
                            if (!wanted.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;

                            string relative = wanted.Substring(prefix.Length).Trim('/');
                            string local = relative.Length == 0
                                ? mount
                                : Path.Combine(mount, relative.Replace('/', Path.DirectorySeparatorChar));

                            if (Directory.Exists(local)) return local;
                        }
                    }
                }
            }
            catch
            {
                // Best effort only - fall back to asking the user.
            }

            return null;
        }

        private static string Normalize(string url) =>
            Uri.UnescapeDataString(url).Replace('\\', '/').TrimEnd('/');
    }
}
