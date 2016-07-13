using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn_Example {
   [ComVisible(true)]
   public class Ribbon : Office.IRibbonExtensibility {
      private const string GroupSmile = "Smileys";
      private const string Smile = "Smile";
      private const string TooltipSmile = "This is a tooltip to smile.";
      private const string HelperSmile = "Don't forget to smile!";
      private const string GroupInfo = "Info";
      private const string Info = "Some information about this Outlook Add-In ...";
      private const string GithubExample = "GitHub-Example";

      private Office.IRibbonUI _ribbon;

      #region IRibbonExtensibility Members

      public string GetCustomUI(string ribbonId) {
         return Resource.Ribbon;
      }

      #endregion

      #region Ribbon Callbacks

      //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

      public void Ribbon_Load(Office.IRibbonUI ribbonUi) {
         _ribbon = ribbonUi;
      }

      public string Group_GetLabel(Office.IRibbonControl control) {
         return GroupSmile;
      }

      public Bitmap Button_GetImage(Office.IRibbonControl control) {
         return Resource.Smile;
      }

      public void Button_Click(Office.IRibbonControl control) {
         MessageBox.Show(Smile);
      }

      public string Button_GetTip(Office.IRibbonControl control) {
         return TooltipSmile;
      }

      public string Button_GetLabel(Office.IRibbonControl control) {
         return Smile;
      }

      public string Group_GetHelperText(Office.IRibbonControl control) {
         return HelperSmile;
      }

      public string GroupInfo_GetLabel(Office.IRibbonControl control) {
         return GroupInfo;
      }

      public string Info_GetLabel(Office.IRibbonControl control) {
         return Info;
      }


      public string TabSmileys_GetLabel(Office.IRibbonControl control) {
         return GroupSmile;
      }

      public string LinkDownload_GetLabel(Office.IRibbonControl control) {
         return GithubExample;
      }

      #endregion
   }
}