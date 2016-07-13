using System;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn_Example {
   public partial class ThisAddIn {
      private void ThisAddIn_Startup(object sender, EventArgs e) {
      }

      private void ThisAddIn_Shutdown(object sender, EventArgs e) {
         // Note: Outlook no longer raises this event. If you have code that 
         //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
      }

      protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject() {
         return new Ribbon();
      }

      #region VSTO generated code

      /// <summary>
      ///    Required method for Designer support - do not modify
      ///    the contents of this method with the code editor.
      /// </summary>
      private void InternalStartup() {
         Startup += ThisAddIn_Startup;
         Shutdown += ThisAddIn_Shutdown;
      }

      #endregion
   }
}