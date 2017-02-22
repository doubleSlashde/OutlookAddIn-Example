namespace OutlookAddIn_Example
{
   using System.Drawing;
   using System.Runtime.InteropServices;
   using System.Windows.Forms;

   using Office = Microsoft.Office.Core;

   [ComVisible(true)]
   public class Ribbon : Office.IRibbonExtensibility
   {
      private Office.IRibbonUI ribbon;

      public string GetCustomUI(string ribbonId)
      {
         return Resource.Ribbon;
      }

      public void Ribbon_Load(Office.IRibbonUI ribbonUi)
      {
         this.ribbon = ribbonUi;
      }

      public string Group_GetLabel(Office.IRibbonControl control)
      {
         return "Hello";
      }

      public Bitmap Button_GetImage(Office.IRibbonControl control)
      {
         return Resource.World;
      }

      public void Button_Click(Office.IRibbonControl control)
      {
         MessageBox.Show("Hello World...");
      }

      public string Button_GetTip(Office.IRibbonControl control)
      {
         return "Hello Tooltip";
      }

      public string Button_GetLabel(Office.IRibbonControl control)
      {
         return "Say hello";
      }

      public string Group_GetHelperText(Office.IRibbonControl control)
      {
         return "Be friendly and say hello!";
      }

      public string GroupInfo_GetLabel(Office.IRibbonControl control)
      {
         return "Info about Hello World";
      }

      public string Info_GetLabel(Office.IRibbonControl control)
      {
         return "Some information about saying hello to the world..";
      }

      public string TabSmileys_GetLabel(Office.IRibbonControl control)
      {
         return "Hello";
      }

      public string LinkDownload_GetLabel(Office.IRibbonControl control)
      {
         return "GitHub-Example";
      }
   }
}