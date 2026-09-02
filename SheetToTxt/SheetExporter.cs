using System;
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
    /// <c>.txt</c> file written next to the workbook, with the workbook's base name.
    /// </summary>
    internal sealed class SheetExporter
    {
        // "Text (Tab delimited) (*.txt)": ANSI (system code page), tab-separated,
        // CRLF line endings, active sheet only. Switch to xlUnicodeText if sheets
        // routinely contain characters outside the system code page.
        private const Excel.XlFileFormat TextFormat = Excel.XlFileFormat.xlTextWindows;

        public ExportResult ExportActiveSheet()
        {
            Excel.Application app = ThisAddIn.Instance?.Application
                ?? throw new InvalidOperationException("The add-in is not initialized.");

            Excel.Workbook workbook = app.ActiveWorkbook
                ?? throw new InvalidOperationException("No workbook is open.");

            WorkbookLocation location = WorkbookLocation.FromWorkbook(workbook);
            if (!location.IsSaved)
            {
                throw new InvalidOperationException(
                    "Save the workbook to a folder first. The .txt is written next to the .xlsx, " +
                    "so the workbook needs a file-system location. (Cloud-only files aren't supported.)");
            }

            if (!(app.ActiveSheet is Excel.Worksheet sheet))
            {
                throw new InvalidOperationException("The active sheet is not a worksheet.");
            }

            string sheetName = sheet.Name;
            string targetPath = location.TargetTextPath;

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
    }
}
