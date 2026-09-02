using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace SheetToTxt
{
    /// <summary>
    /// The on-disk location of a workbook, split into the pieces the exporter needs.
    /// </summary>
    internal sealed class WorkbookLocation
    {
        private WorkbookLocation(bool isSaved, string folder, string baseName, string fullName)
        {
            IsSaved = isSaved;
            Folder = folder;
            BaseName = baseName;
            FullName = fullName;
        }

        /// <summary>False when the workbook has never been saved (no path yet).</summary>
        public bool IsSaved { get; }

        /// <summary>Directory containing the workbook, e.g. <c>C:\Projects\2026</c>.</summary>
        public string Folder { get; }

        /// <summary>Workbook file name without extension, e.g. <c>Budget</c>.</summary>
        public string BaseName { get; }

        /// <summary>Full path of the workbook, or its name if unsaved.</summary>
        public string FullName { get; }

        /// <summary>Full path the .txt export should be written to (workbook folder + base name + ".txt").</summary>
        public string TargetTextPath => Path.Combine(Folder, BaseName + ".txt");

        public static WorkbookLocation FromWorkbook(Excel.Workbook workbook)
        {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            // Excel sets Path to "" until the workbook has been saved at least once.
            // A workbook stored on OneDrive/SharePoint that is synced locally still
            // reports a normal file-system path here; a purely cloud (unsynced) file
            // reports a URL, which Path.Combine/File.WriteAllText cannot use.
            string path = workbook.Path ?? string.Empty;
            bool isSaved = path.Length > 0 && Directory.Exists(path);

            string fileName = workbook.Name ?? "Workbook.xlsx";
            string baseName = Path.GetFileNameWithoutExtension(fileName);
            if (string.IsNullOrEmpty(baseName)) baseName = "Workbook";

            string fullName = isSaved
                ? (workbook.FullName ?? Path.Combine(path, fileName))
                : fileName;

            return new WorkbookLocation(isSaved, path, baseName, fullName);
        }
    }
}
