using Microsoft.Win32;
using System;

namespace Schillings.SwordPhish
{
    public class Configuration
    {
        private static string ROOT_KEY = @"SOFTWARE\Schillings\SwordPhish";

        public static string ReportRecipient => GetValue<string>("ReportRecipient");
        public static string ReportSubject => GetValue<string>("ReportSubject", "SwordPhish Report");
        public static bool DeleteAfterReport => GetValue<int>("DeleteAfterReport", 0) == 1;
        public static bool MoveToJunkAfterReport => GetValue<int>("MoveToJunkAfterReport", 0) == 1;

        private static T GetValue<T>(string value, T defaultValue = default(T))
        {
            T retVal = defaultValue;

            try
            {
                var regView = Environment.Is64BitOperatingSystem && !Environment.Is64BitProcess
                    ? RegistryView.Registry64
                    : RegistryView.Registry32;

                retVal = (T)RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, regView)
                    .OpenSubKey(ROOT_KEY)
                    .GetValue(value, defaultValue);
            }
            catch (Exception e)
            {
            }

            return retVal;
        }
    }
}