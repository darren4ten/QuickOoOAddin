using Microsoft.Office.Tools;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OoOAddin
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors inspectors;
        private Ribbon3 _ribbon = new Ribbon3();
        private OoOSetting settingCtrl;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            settingCtrl = new OoOSetting();
            CustomTaskPane taskPane = this.CustomTaskPanes.Add(settingCtrl, "OoO Setting");
            settingCtrl.SetPanel(this.Application,taskPane);
            _ribbon.SetApplication(this.Application, taskPane);
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return _ribbon;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
