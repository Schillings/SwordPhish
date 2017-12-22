using Microsoft.Office.Interop.Outlook;
using System;

namespace Schillings.SwordPhish.Shared
{
    public static class Helpers
    {
        public static void SetUserProperty<T>(MailItem currentMailItem, string propName, T propValue)
        {
            var prop = currentMailItem.UserProperties.Find(propName);

            if (prop != null)
            {
                prop.Value = propValue.ToString();
            }
            else
            {
                prop = currentMailItem.UserProperties.Add(propName, OlUserPropertyType.olText);
                prop.Value = propValue.ToString();
            }
        }

        public static T GetUserProperty<T>(MailItem currentMailItem, string propName)
        {
            var prop = currentMailItem.UserProperties.Find(propName);

            if (prop != null)
            {
                return Convert.ChangeType(prop.Value, typeof(T));
            }

            return default(T);
        }
    }
}