# SwordPhish
**At some point technology will fail and your users will be the last line of defence. SwordPhish allows users to easily report suspicious e-mails to your IT and security teams.**

![SwordPhish](http://g.recordit.co/DEALRp32ml.gif)

SwordPhish is a very simple button that sits within the users Outlook toolbar. One click and the suspicious e-mail is instantly reported and contains all metadata required for investigation.

**SwordPhish should be underpinned by your security awareness programme**. Suspicious e-mails can't be reported if your users don't know what they are looking for or what to deem suspicious.

The recipient of reports receives the original reported e-mail as an  attachment along with every header to aid further investigation.

![SwordPhish Report](https://i.imgur.com/43ZC625.png)

# Requirements
SwordPhish requires Microsoft Office 2007+, .NET Framework 4.0+ and [VTSO 4.0+](https://www.microsoft.com/en-us/download/details.aspx?id=48217) to be installed, and works across both X86 and X64 platforms.

# Installation
The easiest way to install SwordPhish is via the pre-compiled MSI installers under the [Releases tab](https://github.com/Schillings/SwordPhish/releases) - make sure you run as an administrator.

If you are so inclined you can build SwordPhish from source. Simply clone this repository and compile with Visual Studio 2010+.

# Deployment
You can deploy SwordPhish as you would any other application: manually, Group Policy, batch file at logon, or your favourite systems management solution.

# Options
SwordPhish just needs to know where to send reports to. You can set other options via the MSI installer or property flags:

Flag | Description
---- | -----------
RECIPIENTPROPERTY | E-mail address where to send SwordPhish reports
SUBJECTPROPERTY | Subject of SwordPhish reports
ACTIONPROPERTY | What to do after a user reports an e-mail. 0 = Just send the report, 1 = Report and move the e-mail to "Junk", 2 = Report and delete the e-mail from the user's Inbox.

For example: `msiexec /i Schillings.SwordPhish.Installer.msi RECIPIENTPROPERTY="reports@mysoc.com" ACTIONPROPERTY=2`

To make your life easier for reporting it is recommended to set the recipient address to a ticketing system.

# License and Disclaimer
SwordPhish is [licensed under Apache 2.0](LICENSE).

In no event and under no legal theory, whether in tort (including negligence), contract, or otherwise, unless required by applicable law (such as deliberate and grossly negligent acts) or agreed to in writing, shall any Contributor be liable to You for damages, including any direct, indirect, special, incidental, or consequential damages of any character arising as a result of the use or inability to use the Work (including but not limited to damages for loss of goodwill, work stoppage, computer failure or malfunction, or any and all other commercial damages or losses), even if such Contributor has been advised of the possibility of such damages.