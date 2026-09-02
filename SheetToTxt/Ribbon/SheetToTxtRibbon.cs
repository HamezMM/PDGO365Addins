using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace SheetToTxt.Ribbon
{
    /// <summary>
    /// XML ribbon for the add-in. Callbacks are kept thin: they gather context,
    /// delegate to <see cref="SheetExporter"/>, and report the result. No exception
    /// is allowed to escape a callback (Office can disable the add-in if one does).
    /// </summary>
    [ComVisible(true)]
    public class SheetToTxtRibbon : IRibbonExtensibility
    {
        private const string Title = "Sheet to TXT";
        private const string RibbonResourceName = "SheetToTxt.Ribbon.SheetToTxtRibbon.xml";

        private IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID) => GetResourceText(RibbonResourceName);

        public void OnLoad(IRibbonUI ribbonUI) => _ribbon = ribbonUI;

        /// <summary>Enabled only when a worksheet is active.</summary>
        public bool ExportSheetButton_GetEnabled(IRibbonControl control)
        {
            try
            {
                Excel.Application app = ThisAddIn.Instance?.Application;
                return app?.ActiveWorkbook != null && app.ActiveSheet is Excel.Worksheet;
            }
            catch
            {
                return false;
            }
        }

        public void ExportSheetButton_OnAction(IRibbonControl control)
        {
            try
            {
                ExportResult result = new SheetExporter().ExportActiveSheet();
                MessageBox.Show(
                    $"Exported sheet \"{result.SheetName}\" to:{Environment.NewLine}{result.TargetPath}",
                    Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            using (Stream stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException(
                        $"Embedded ribbon resource '{resourceName}' was not found. " +
                        "Check the file's Build Action (Embedded Resource) and LogicalName in the .csproj.");
                }

                using (StreamReader reader = new StreamReader(stream))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}
