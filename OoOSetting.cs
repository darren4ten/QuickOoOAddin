using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools;

namespace OoOAddin
{
    public partial class OoOSetting : UserControl
    {
        private CustomTaskPane _pane;
        private Microsoft.Office.Interop.Outlook.Application _app;
        public OoOSetting()
        {
            InitializeComponent();
        }

        public void SetPanel(Microsoft.Office.Interop.Outlook.Application application, CustomTaskPane pane)
        {
            _pane = pane;
            _pane.Width = 500;
            _app = application;

            LoadControl();
        }

        private void LoadControl()
        {
            txtSubject_OoO.Text = Settings1.Default.txtSubject_OoO;
            txtReceivers_OoO.Text = Settings1.Default.txtReceivers_OoO;
            txtReceiversRequired_OoO.Text = Settings1.Default.txtReceiversRequired_OoO;
            txtBody_OoO.Text = Settings1.Default.txtBody_OoO;
            txtSubject_WFH.Text = Settings1.Default.txtSubject_WFH;
            txtReceivers_WFH.Text = Settings1.Default.txtReceivers_WFH;
            txtReceiversRequired_WFH.Text = Settings1.Default.txtReceiversRequired_WFH;
            txtBody_WFH.Text = Settings1.Default.txtBody_WFH;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                var ps = Properties.Settings.Default.Properties;
                Settings1.Default.txtSubject_OoO = txtSubject_OoO.Text;
                Settings1.Default.txtReceivers_OoO = txtReceivers_OoO.Text;
                Settings1.Default.txtBody_OoO = txtBody_OoO.Text;
                Settings1.Default.txtSubject_WFH = txtSubject_WFH.Text;
                Settings1.Default.txtReceivers_WFH = txtReceivers_WFH.Text;
                Settings1.Default.txtBody_WFH = txtBody_WFH.Text;
                Settings1.Default.txtReceiversRequired_OoO = txtReceiversRequired_OoO.Text;
                Settings1.Default.txtReceiversRequired_WFH = txtReceiversRequired_WFH.Text;
                Settings1.Default.Save();
                lblSaveResult.Text = "Save success!";
                lblSaveResult.ForeColor = Color.MediumSeaGreen;
                lblSaveResult.Update();
                Thread.Sleep(1000);
                _pane.Visible = false;
                lblSaveResult.Text = "";
            }
            catch (Exception exception)
            {
                lblSaveResult.Text = "Save failed!" + exception.Message + exception.StackTrace;
                lblSaveResult.ForeColor = Color.Red;
            }
         
        }

        private void button1_Click(object sender, EventArgs e)
        {
            _pane.Visible = false;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
