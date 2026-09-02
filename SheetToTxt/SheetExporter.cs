using System;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SheetToTxt
{
    /// <summary>Outcome of a successful export.</summary>
    internal sealed class ExportResult
    {
        public ExportResult(string sheetName, string targetPath)
        {
            SheetName = sheetName;
            TargetPath = targetPath;
        }

        public string SheetName { get; }
        public string TargetPath { get; }
    }

    /// <summary>
    /// Exports the active worksheet of the active workbook as a tab-delimited
    /// <c>.txt</c> file. For a normal or locally-synced workbook the file is written
    /// next to it with the workbook's base name; for a cloud-only or never-saved
    /// workbook the user is asked where to write it.
    /// </summary>
    internal sealed class SheetExporter
    {
        // "Text (Tab delimited) (*.txt)": ANSI (system code page), tab-separated,
        // CRLF line endings, active sheet only. Switch to xlUnicodeText if sheets
        // routinely contain characters outside the system code page.
        private const Excel.XlFileFormat TextFormat = Excel.XlFileFormat.xlTextWindows;

        /// <summary>Runs the export. Returns null if the user cancels the "save as" prompt.</summary>
        public ExportResult ExportActiveSheet()
        {
            Excel.Application app = ThisAddIn.Instance?.Application
                ?? throw new InvalidOperationException("The add-in is not initialized.");

            Excel.Workbook workbook = app.ActiveWorkbook
                ?? throw new InvalidOperationException("No workbook is open.");

            if (!(app.ActiveSheet is Excel.Worksheet sheet))
            {
                throw new InvalidOperationException("The active sheet is not a worksheet.");
            }

            WorkbookLocation location = WorkbookLocation.FromWorkbook(workbook);

            string targetPath;
            if (location.HasLocalFolder)
            {
                // Normal file, or a OneDrive/SharePoint workbook synced to this machine:
                // write next to the workbook, no prompt (unchanged behaviour).
                targetPath = location.TargetTextPath;
            }
            else
            {
                // Cloud-only workbook (opened from SharePoint/OneDrive without a local
                // sync) or one that has never been saved: there is no folder to write
                // next to, so ask where the .txt should go.
                targetPath = PromptForTargetPath(app, location.BaseName);
                if (targetPath == null) return null; // user cancelled
            }

            string sheetName = sheet.Name;

            bool screenUpdating = app.ScreenUpdating;
            bool displayAlerts = app.DisplayAlerts;
            app.ScreenUpdating = false;
            app.DisplayAlerts = false; // overwrite prompt + "features may be lost" prompt
            try
            {
                // Copy the active sheet into a new single-sheet workbook, then let Excel
                // write the text file. This is exactly File > Save As > Text (Tab delimited),
                // so number/date/formula formatting matches what the user sees on screen.
                sheet.Copy(Type.Missing, Type.Missing);
                Excel.Workbook exportBook = app.ActiveWorkbook;
                try
                {
                    exportBook.SaveAs(targetPath, TextFormat);
                }
                finally
                {
                    exportBook.Close(SaveChanges: false);
                }

                // Return focus to the workbook the user was on.
                workbook.Activate();
            }
            finally
            {
                app.DisplayAlerts = displayAlerts;
                app.ScreenUpdating = screenUpdating;
            }

            return new ExportResult(sheetName, targetPath);
        }

        private static string PromptForTargetPath(Excel.Application app, string baseName)
        {
            string initialDir =
                Environment.GetEnvironmentVariable("OneDriveCommercial")
                ?? Environment.GetEnvironmentVariable("OneDrive")
                ?? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (string.IsNullOrEmpty(initialDir) || !Directory.Exists(initialDir))
                initialDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            using (var dialog = new SaveFileDialog
            {
                Title = "Export sheet to .txt",
                Filter = "Text (Tab delimited) (*.txt)|*.txt|All files (*.*)|*.*",
                DefaultExt = "txt",
                AddExtension = true,
                OverwritePrompt = true,
                FileName = baseName + ".txt",
                InitialDirectory = initialDir,
            })
            {
                DialogResult result = dialog.ShowDialog(new WindowHandle((IntPtr)app.Hwnd));
                return result == DialogResult.OK ? dialog.FileName : null;
            }
        }

        /// <summary>Wraps Excel's top-level window so dialogs are modal to Excel.</summary>
        private sealed class WindowHandle : IWin32Window
        {
            public WindowHandle(IntPtr handle) => Handle = handle;
            public IntPtr Handle { get; }
        }
    }
}
