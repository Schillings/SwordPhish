using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Schillings.SwordPhish.Shared;
using Schillings.SwordPhish.Shared.Properties;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace Schillings.SwordPhish
{
    [ComVisible(true)]
    public class Ribbon : IRibbonExtensibility
    {
        private IRibbonUI ribbon;

        public Ribbon()
        {
        }

        public void ReportMailItem(IRibbonControl control)
        {
            Globals.ThisAddIn.SendReport(GetMailItemBasedOnControl(control));
            ribbon.InvalidateControl(control.Id);
        }

        public bool GetButtonEnabled(IRibbonControl control)
        {
            MailItem mailItem = GetMailItemBasedOnControl(control);

            return Globals.ThisAddIn.HasConfig
                && mailItem != null
                && !Helpers.GetUserProperty<bool>(mailItem, Constants.EMAIL_REPORTED);
        }

        public string GetButtonLabel(IRibbonControl control)
        {
            MailItem mailItem = GetMailItemBasedOnControl(control);

            if (mailItem == null)
                return Resources.ReportButton;

            return Helpers.GetUserProperty<bool>(mailItem, Constants.EMAIL_REPORTED)
                ? Resources.Reported
                : Resources.ReportButton;
        }

        public string GetButtonSuperTip(IRibbonControl control)
        {
            var mailItem = GetMailItemBasedOnControl(control);

            if (mailItem == null)
                return Resources.ReportButtonHelp;

            return Helpers.GetUserProperty<bool>(mailItem, Constants.EMAIL_REPORTED)
                ? String.Format(Resources.ReportedDate, Helpers.GetUserProperty<string>(mailItem, Constants.EMAIL_REPORTED_DATE))
                : Resources.ReportButtonHelp;
        }

        public Bitmap GetButtonImage(IRibbonControl control)
        {
            var mailItem = GetMailItemBasedOnControl(control);

            if (mailItem == null)
                return Resources.icon;

            return Helpers.GetUserProperty<bool>(GetMailItemBasedOnControl(control), Constants.EMAIL_REPORTED)
                ? Resources.icon_reported
                : Resources.icon;
        }

        public String GetGroupLabel(IRibbonControl control)
        {
            return Resources.GroupLabel;
        }

        private MailItem GetMailItemBasedOnControl(IRibbonControl control)
        {
            Object selectedObjected = null;

            if (control.Id.Equals("buttonMailReport"))
                selectedObjected = Globals.ThisAddIn.Application.ActiveExplorer().Selection[1];
            else if (control.Id.Equals("buttonReadReport"))
                selectedObjected = Globals.ThisAddIn.Application.ActiveInspector().CurrentItem;

            if (selectedObjected != null && selectedObjected is MailItem)
                return selectedObjected as MailItem;

            return null;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Schillings.SwordPhish.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            Globals.ThisAddIn.Ribbon = ribbonUI;
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
