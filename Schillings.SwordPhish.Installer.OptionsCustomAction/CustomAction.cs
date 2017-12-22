using Microsoft.Deployment.WindowsInstaller;
using Microsoft.Win32;
using System;
using System.IO;

namespace Schillings.SwordPhish.Installer.OptionsCustomAction
{
    public class CustomActions
    {
        private static string ROOT_KEY = @"SOFTWARE\Schillings\SwordPhish";

        [CustomAction]
        public static ActionResult SaveOptions(Session session)
        {
            System.Diagnostics.Debugger.Launch();

            var recipient = session["RECIPIENTPROPERTY"];
            var subject = session["SUBJECTPROPERTY"];
            var action = session["ACTIONPROPERTY"];

            if (String.IsNullOrWhiteSpace(recipient) || String.IsNullOrWhiteSpace(subject))
                return ActionResult.Failure;

            try
            {
                var regView = Environment.Is64BitOperatingSystem && !Environment.Is64BitProcess
                    ? RegistryView.Registry64
                    : RegistryView.Registry32;

                var subKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, regView).CreateSubKey(ROOT_KEY);
                subKey.SetValue("ReportRecipient", recipient, RegistryValueKind.String);
                subKey.SetValue("ReportSubject", subject, RegistryValueKind.String);

                subKey.SetValue("MoveToJunkAfterReport", action.Equals("1") ? 1 : 0, RegistryValueKind.DWord);
                subKey.SetValue("DeleteAfterReport", action.Equals("2") ? 1 : 0, RegistryValueKind.DWord);

                return ActionResult.Success;
            }
            catch (Exception e)
            {
                return ActionResult.Failure;
            }
        }
    }
}