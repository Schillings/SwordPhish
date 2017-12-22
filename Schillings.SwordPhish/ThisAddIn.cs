using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Schillings.SwordPhish.Shared;
using Schillings.SwordPhish.Shared.Properties;
using System;
using System.Text;

namespace Schillings.SwordPhish
{
    public partial class ThisAddIn
    {
        public IRibbonUI Ribbon;
        public bool HasConfig = false;

        private Explorers _explorers;
        private Explorer _currentExplorer;
        private Inspectors _inspectors;

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            HasConfig = !String.IsNullOrWhiteSpace(Configuration.ReportRecipient);

            _explorers = Application.Explorers;
            _explorers.NewExplorer += new ExplorersEvents_NewExplorerEventHandler(NewExplorer);

            _inspectors = Application.Inspectors;
            _inspectors.NewInspector += new InspectorsEvents_NewInspectorEventHandler(NewInspector);

            _currentExplorer = this.Application.ActiveExplorer();
            _currentExplorer.SelectionChange += new ExplorerEvents_10_SelectionChangeEventHandler(Explorer_SelectionChange);
        }

        private void Application_ItemLoad(object item)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void NewExplorer(Explorer explorer)
        {
            explorer.SelectionChange += new ExplorerEvents_10_SelectionChangeEventHandler(Explorer_SelectionChange);
        }

        private void NewInspector(Inspector inspector)
        {
            Object currentObject = inspector.CurrentItem;

            if (Ribbon != null)
                Ribbon.InvalidateControl("buttonReadReport");
        }

        private void Explorer_SelectionChange()
        {
            if (Application.ActiveExplorer().Selection.Count > 0)
            {
                Object selectedObject = Application.ActiveExplorer().Selection[1];
                
                if (Ribbon != null)
                    Ribbon.InvalidateControl("buttonMailReport");
            }
        }

        public void SendReport(MailItem reportedMailItem)
        {
            Helpers.SetUserProperty(reportedMailItem, Constants.EMAIL_REPORTED, true);
            Helpers.SetUserProperty(reportedMailItem, Constants.EMAIL_REPORTED_DATE, DateTime.Now);
            reportedMailItem.Save();

            CreateAndSendReport(reportedMailItem);

            if (Configuration.DeleteAfterReport)
            {
                reportedMailItem.Delete();
            }
            else if (Configuration.MoveToJunkAfterReport)
            {
                MAPIFolder junkFolder = Application.ActiveExplorer().Session.GetDefaultFolder(OlDefaultFolders.olFolderJunk);
                reportedMailItem.Move(junkFolder);
            }
        }

        private void CreateAndSendReport(MailItem reportedMailItem)
        {
            var supportMail = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
            supportMail.To = Configuration.ReportRecipient;
            supportMail.Subject = Configuration.ReportSubject;
            supportMail.Attachments.Add(reportedMailItem, OlAttachmentType.olByValue, 1, reportedMailItem.Subject);
            supportMail.BodyFormat = OlBodyFormat.olFormatHTML;
            supportMail.HTMLBody = GenerateReportBody(reportedMailItem);
            supportMail.Importance = OlImportance.olImportanceHigh;
            supportMail.Send();
        }

        private string GenerateReportBody(MailItem reportedMailItem)
        {
            var emailHeadersTable = new StringBuilder();

            foreach (var header in reportedMailItem.Headers())
            {
                foreach (var value in header)
                {
                    emailHeadersTable.AppendFormat(Resources.EmailHeaderTableRowHtml, header.Key, value);
                }
            }

            var body = Encoding.UTF8.GetString(Resources.email);
            body = body.Replace(String.Format("{0}{1}{0}", Constants.EMAIL_TOKEN_SEPERATOR, Constants.EMAIL_HEADERS_TOKEN), emailHeadersTable.ToString());

            return body;
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