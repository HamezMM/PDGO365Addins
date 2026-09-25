using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SpeakerRename.UI;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace SpeakerRename.Ribbon
{
    /// <summary>
    /// XML ribbon for the add-in. Callbacks are kept thin: they gather context,
    /// delegate to <see cref="SpeakerFieldScanner"/> / <see cref="SpeakerFieldReplacer"/>,
    /// and report the result. No exception is allowed to escape a callback (Office can
    /// disable the add-in if one does).
    /// </summary>
    [ComVisible(true)]
    public class SpeakerRenameRibbon : IRibbonExtensibility
    {
        private const string Title = "Speaker Rename";
        private const string RibbonResourceName = "SpeakerRename.Ribbon.SpeakerRenameRibbon.xml";

        private IRibbonUI _ribbon;

        public string GetCustomUI(string ribbonID) => GetResourceText(RibbonResourceName);

        public void OnLoad(IRibbonUI ribbonUI) => _ribbon = ribbonUI;

        /// <summary>
        /// Re-runs <see cref="RenameSpeakersButton_GetEnabled"/>. Called by
        /// <see cref="ThisAddIn"/> on document changes.
        /// </summary>
        public void InvalidateRenameButton()
        {
            try { _ribbon?.InvalidateControl("RenameSpeakersButton"); }
            catch { /* ribbon not ready yet; the next event will catch up */ }
        }

        /// <summary>Enabled only when a document is open.</summary>
        public bool RenameSpeakersButton_GetEnabled(IRibbonControl control)
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

        public void RenameSpeakersButton_OnAction(IRibbonControl control)
        {
            try
            {
                Word.Application app = ThisAddIn.Instance.Application;
                if (app.Documents.Count == 0)
                    throw new InvalidOperationException("Open a document first.");

                Word.Document doc = app.ActiveDocument;
                List<SpeakerField> fields = SpeakerFieldScanner.Scan(doc);
                if (fields.Count == 0)
                {
                    MessageBox.Show("No \"Speaker N\" labels were found in this document.",
                        Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                using (var form = new RenameSpeakersForm(fields, doc.Name))
                {
                    if (form.ShowDialog(WordWindow(app)) != DialogResult.OK) return;
                }

                RenameResult result = SpeakerFieldReplacer.Replace(app, doc, fields);
                MessageBox.Show(Summarize(result), Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private static string Summarize(RenameResult result)
        {
            string summary =
                $"Replaced {result.OccurrencesReplaced} occurrence{(result.OccurrencesReplaced == 1 ? "" : "s")} " +
                $"of {result.SpeakersRenamed} speaker{(result.SpeakersRenamed == 1 ? "" : "s")}.";
            if (result.SpeakersSkipped > 0)
                summary += $"{Environment.NewLine}{result.SpeakersSkipped} speaker{(result.SpeakersSkipped == 1 ? " was" : "s were")} left blank and not changed.";
            return summary + $"{Environment.NewLine}{Environment.NewLine}Press Ctrl+Z once to undo the whole rename.";
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
