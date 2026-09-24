using System;
using Microsoft.Office.Core;

namespace BracketFields
{
    public partial class ThisAddIn
    {
        /// <summary>Set during startup so feature code can reach the add-in without <c>Globals</c>.</summary>
        internal static ThisAddIn Instance { get; private set; }

        /// <summary>The ribbon instance, kept so document events can re-run its <c>getEnabled</c> callbacks.</summary>
        internal Ribbon.BracketFieldsRibbon Ribbon { get; private set; }

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Instance = this;

            // getEnabled runs only when the ribbon is built or the control is
            // invalidated. DocumentChange fires when a document is created, opened,
            // closed, or switched to, which is exactly when "is a document open?" changes.
            Application.DocumentChange += InvalidateFillButton;

            // Keep this fast. Slow work here delays Word's launch and can get the
            // add-in disabled by Office. Defer real work to the first ribbon action.
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Instance = null;
        }

        private void InvalidateFillButton() => Ribbon?.InvalidateFillButton();

        /// <summary>Registers the XML ribbon defined in <see cref="Ribbon.BracketFieldsRibbon"/>.</summary>
        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Ribbon = new Ribbon.BracketFieldsRibbon();
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
