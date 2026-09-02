using System;
using Microsoft.Office.Core;

namespace SheetToTxt
{
    public partial class ThisAddIn
    {
        /// <summary>Set during startup so feature code can reach the add-in without <c>Globals</c>.</summary>
        internal static ThisAddIn Instance { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Instance = this;
            // Keep this fast. Slow work here delays Excel's launch and can get the
            // add-in disabled by Office. Defer real work to the first ribbon action.
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Instance = null;
        }

        /// <summary>Registers the XML ribbon defined in <see cref="Ribbon.SheetToTxtRibbon"/>.</summary>
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon.SheetToTxtRibbon();
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
