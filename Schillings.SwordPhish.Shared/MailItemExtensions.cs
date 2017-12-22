using Microsoft.Office.Interop.Outlook;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace Schillings.SwordPhish.Shared
{
    public static class MailItemExtensions
    {
        private const string HeaderRegex = @"^(?<header_key>[-A-Za-z0-9]+)(?<seperator>:[ \t]*)(?<header_value>([^\r\n]|\r\n[ \t]+)*)(?<terminator>\r\n)";
        private const string TransportMessageHeadersSchema = "http://schemas.microsoft.com/mapi/proptag/0x007D001E";

        public static IList<string> Headers(this MailItem mailItem, string name)
        {
            var headers = mailItem.Headers();
            if (headers.Contains(name))
                return headers[name].ToList();

            return new List<string>();
        }

        public static ILookup<string, string> Headers(this MailItem mailItem)
        {
            var headerString = mailItem.HeaderString();
            var headerMatches = Regex.Matches(headerString, HeaderRegex, RegexOptions.Multiline).Cast<Match>();

            return headerMatches.ToLookup(
                h => h.Groups["header_key"].Value,
                h => h.Groups["header_value"].Value);
        }

        public static string GetHash(this MailItem mailItem)
        {
            var hash = (mailItem.Subject.GetHashCode() + mailItem.SenderEmailAddress.GetHashCode()).ToString();
            var md5 = new MD5CryptoServiceProvider();
            md5.ComputeHash(ASCIIEncoding.ASCII.GetBytes(hash));
 
            var strBuilder = new StringBuilder();
            for (var i = 0; i < md5.Hash.Length; i++)
            {
                strBuilder.Append(md5.Hash[i].ToString("x2"));
            }

            return strBuilder.ToString();
        }

        private static string HeaderString(this MailItem mailItem)
        {
            return (string)mailItem.PropertyAccessor
                .GetProperty(TransportMessageHeadersSchema);
        }
    }
}