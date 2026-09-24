using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using BracketFields.UI;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace BracketFields.Ribbon
{
    /// <summary>
    /// XML ribbon for the add-in. Callbacks are kept thin: they gather context,
    /// delegate to <see cref="BracketFieldScanner"/> / <see cref="BracketFieldReplacer"/>,
    /// and report the result. No exception is allowed to escape a callback (Office can
    /// disable the add-in if one does).
    /// </summary>
    [ComVisible(true)]
    public class BracketFieldsRibbon : IRibbonExtensibility
    {
        private const string Title = "Bracket Fields";
        private const string RibbonResourceName = "BracketFields.Ribbon.BracketFieldsRibbon.xml";

        private IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID) => GetResourceText(RibbonResourceName);

        public void OnLoad(IRibbonUI ribbonUI) => _ribbon = ribbonUI;

        /// <summary>
        /// Re-runs <see cref="FillFieldsButton_GetEnabled"/>. Called by
        /// <see cref="ThisAddIn"/> on document changes.
        /// </summary>
        public void InvalidateFillButton()
        {
            try { _ribbon?.InvalidateControl("FillFieldsButton"); }
            catch { /* ribbon not ready yet; the next event will catch up */ }
        }

        /// <summary>Enabled only when a document is open.</summary>
        public bool FillFieldsButton_GetEnabled(IRibbonControl control)
        {
            try
            {
                Word.Application app = ThisAddIn.Instance?.Application;
                return app != null && app.Documents.Count > 0;
            }
            catch
            {
                return false;
            }
        }

        public void FillFieldsButton_OnAction(IRibbonControl control)
        {
            try
            {
                Word.Application app = ThisAddIn.Instance.Application;
                if (app.Documents.Count == 0)
                    throw new InvalidOperationException("Open a document first.");

                Word.Document doc = app.ActiveDocument;
                List<BracketField> fields = BracketFieldScanner.Scan(doc);
                if (fields.Count == 0)
                {
                    MessageBox.Show("No [bracketed] fields were found in this document.",
                        Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (var form = new FillFieldsForm(fields, doc.Name))
                {
                    if (form.ShowDialog(WordWindow(app)) != DialogResult.OK) return;
                }

                FillResult result = BracketFieldReplacer.Replace(app, doc, fields);
                MessageBox.Show(Summarize(result), Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static string Summarize(FillResult result)
        {
            string summary =
                $"Replaced {result.OccurrencesReplaced} occurrence{(result.OccurrencesReplaced == 1 ? "" : "s")} " +
                $"of {result.FieldsFilled} field{(result.FieldsFilled == 1 ? "" : "s")}.";
            if (result.FieldsSkipped > 0)
                summary += $"{Environment.NewLine}{result.FieldsSkipped} field{(result.FieldsSkipped == 1 ? " was" : "s were")} left blank and not changed.";
            return summary + $"{Environment.NewLine}{Environment.NewLine}Press Ctrl+Z once to undo the whole fill.";
        }

        /// <summary>Word's active window, so the dialog is owned by (and centred on) Word.</summary>
        private static IWin32Window WordWindow(Word.Application app)
        {
            try
            {
                Word.Window window = app.ActiveWindow;
                return new WindowHandle(new IntPtr(window.Hwnd));
            }
            catch
            {
                return null;
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
