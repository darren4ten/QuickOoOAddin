using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using Application = Microsoft.Office.Interop.Outlook.Application;
using Office = Microsoft.Office.Core;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon3();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace OoOAddin
{
    [ComVisible(true)]
    public class Ribbon3 : Office.IRibbonExtensibility
    {
        private Application _app;
        private Office.IRibbonUI ribbon;
        private CustomTaskPane _settignCtrl;

        public Ribbon3()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OoOAddin.Ribbon3.xml");
        }

        #endregion

        public void SetApplication(Microsoft.Office.Interop.Outlook.Application app, CustomTaskPane ctrl)
        {
            _app = app;
            _settignCtrl = ctrl;
        }

        private Outlook.ExchangeUser GetCurrentUserInfo()
        {
            Outlook.AddressEntry addrEntry = _app.Session.CurrentUser.AddressEntry;
            if (addrEntry.Type == "EX")
            {
                Outlook.ExchangeUser currentUser =
                    _app.Session.CurrentUser.AddressEntry.GetExchangeUser();
                if (currentUser != null)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine("Name: "
                                  + currentUser.Name);
                    sb.AppendLine("STMP address: "
                                  + currentUser.PrimarySmtpAddress);
                    sb.AppendLine("Title: "
                                  + currentUser.JobTitle);
                    sb.AppendLine("Department: "
                                  + currentUser.Department);
                    sb.AppendLine("Location: "
                                  + currentUser.OfficeLocation);
                    sb.AppendLine("Business phone: "
                                  + currentUser.BusinessTelephoneNumber);
                    sb.AppendLine("Mobile phone: "
                                  + currentUser.MobileTelephoneNumber);
                }

                return currentUser;
            }

            return null;
        }

        public void OnBtnOoO(object sender, RibbonControlEventArgs e = null)
        {
            BuildEmail(MailType.OoO);
        }

        public void OnBtnWFH(object sender)
        {
            BuildEmail(MailType.WFH);
        }

        public void OnConfigTemplates(object sender)
        {
            _settignCtrl.Visible = true;
        }

        #region Ribbon Callbacks

        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Business

        enum MailType
        {
            OoO,
            WFH
        }


        private void BuildEmail(MailType mailType)
        {
            Microsoft.Office.Interop.Outlook.AppointmentItem agendaMeeting =
                _app.Application.CreateItem(Outlook.OlItemType.olAppointmentItem);
            agendaMeeting.MeetingStatus =
                Microsoft.Office.Interop.Outlook.OlMeetingStatus.olMeeting;
            agendaMeeting.Location = "";
            agendaMeeting.Importance = Outlook.OlImportance.olImportanceLow;
            agendaMeeting.AllDayEvent = true;
            var user = GetCurrentUserInfo();
            if (mailType == MailType.OoO)
            {
                agendaMeeting.Subject = GetHandledSubject(Settings1.Default.txtSubject_OoO, user.Name);
                //agendaMeeting.Body = Settings1.Default.txtBody_OoO;
                SetupRecipients(agendaMeeting.Recipients, Settings1.Default.txtReceivers_OoO);
                SetupBody(agendaMeeting, Settings1.Default.txtBody_OoO);
            }
            else
            {
                agendaMeeting.Subject = GetHandledSubject(Settings1.Default.txtSubject_WFH, user.Name);
                //agendaMeeting.Body = Settings1.Default.txtBody_WFH;
                SetupRecipients(agendaMeeting.Recipients, Settings1.Default.txtReceivers_WFH);
                SetupBody(agendaMeeting, Settings1.Default.txtBody_WFH);
            }

            agendaMeeting.BusyStatus = Outlook.OlBusyStatus.olFree;

            agendaMeeting.Display(false);
        }

        private void SetupBody(Outlook.AppointmentItem agendaMeeting, string body)
        {
            //if (IsHtmlBody(body))
            //{
            //    agendaMeeting.RTFBody = body;
            //}

            agendaMeeting.Body = body;
        }

        private void SetupRecipients(Outlook.Recipients recipients, string recipientStr)
        {
            GetHandledRecipients(recipientStr)?.ForEach(r =>
            {
                Outlook.Recipient recipient =
                    recipients.Add(r);
                recipient.Type =
                    (int)Outlook.OlMeetingRecipientType.olOptional;
            });
        }

        private List<string> GetHandledRecipients(string recipients)
        {
            return recipients.Split(';').ToList();
        }

        private string GetHandledSubject(string subject, string userName)
        {
            return subject?.Replace("$name$", userName);
        }

        private bool IsHtmlBody(string body)
        {
            if (!string.IsNullOrEmpty(body) && body.StartsWith("<html>"))
            {
                return true;
            }
            return false;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
