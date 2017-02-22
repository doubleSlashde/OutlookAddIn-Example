namespace OutlookAddIn_Example
{
   using System;

   using Office = Microsoft.Office.Core;

   public partial class ThisAddIn
   {
      protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
      {
         return new Ribbon();
      }

      private void ThisAddIn_Startup(object sender, EventArgs e)
      {
      }

      private void ThisAddIn_Shutdown(object sender, EventArgs e)
      {
         // Note: Outlook no longer raises this event. If you have code that 
         // must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
      }

      /// <summary>
      ///    Required method for Designer support - do not modify
      ///    the contents of this method with the code editor.
      /// </summary>
      private void InternalStartup()
      {
         this.Startup += this.ThisAddIn_Startup;
         this.Shutdown += this.ThisAddIn_Shutdown;
      }
   }
}