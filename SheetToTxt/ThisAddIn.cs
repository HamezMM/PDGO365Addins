using System;
using Microsoft.Office.Core;

namespace SheetToTxt
{
    public partial class ThisAddIn
    {
        /// <summary>Set during startup so feature code can reach the add-in without <c>Globals</c>.</summary>
        internal static ThisAddIn Instance { get; private set; }

        /// <summary>The ribbon instance, kept so workbook/sheet events can re-run its <c>getEnabled</c> callbacks.</summary>
        internal Ribbon.SheetToTxtRibbon Ribbon { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Instance = this;

            // getEnabled runs only when the ribbon is built or the control is
            // invalidated. Without these hooks the Export button keeps whatever
            // state it had at load time (disabled, if Excel opened with no
            // workbook) even after a workbook is opened.
            // WorkbookActivate covers newly created, opened, and switched-to workbooks.
            Application.WorkbookActivate += _ => InvalidateExportButton();
            Application.WorkbookDeactivate += _ => InvalidateExportButton();
            Application.SheetActivate += _ => InvalidateExportButton();

            // Keep this fast. Slow work here delays Excel's launch and can get the
            // add-in disabled by Office. Defer real work to the first ribbon action.
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Instance = null;
        }

        private void InvalidateExportButton() => Ribbon?.InvalidateExportButton();

        /// <summary>Registers the XML ribbon defined in <see cref="Ribbon.SheetToTxtRibbon"/>.</summary>
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Ribbon = new Ribbon.SheetToTxtRibbon();
            return Ribbon;
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
