using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;

namespace GoldMS2RDMSExceptRpt
{
    class Mail
    {
        private string getMsg()
        {
            string msgFile = "";
            try
            {
                msgFile = Properties.Settings.Default.emailMsg;
                if (!File.Exists(msgFile))
                {
                    // try MyDocuments folder of the current user
                    msgFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\" + msgFile ;
                }
            }
            catch (ConfigurationException e)
            {
                Logger.Warning(e.Message,"Program");
                msgFile = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\GoldMS2RDMSExceptRpt.htm" ;
            }
            if (!File.Exists(msgFile))
            {
                return "<html>" +
                    "<head>" +
                    "<title>GOLD--MS2--RDMS System Compare Exception Report</title>" +
                    "</head>" +
                    "<body style='background-color:#E6E6E6;'>" +
                    "<div style='font-family: Georgia, Arial; font-size:14px; '>Team,<br /><br />" +
                    "<br />" +
                    "The GOLD-MS2-RDMS System Compare Exception report has been updated and can now be retrieved via its link on the right side of the <a href='https://vtr2b.web.boeing.com/ids/escmtools/default.aspx'>ESCM Team SharePoint</a><br /><br /><br /><br /><br />" +
                    "<br />" +
                    "</div>" +
                    "</body>" +
                    "</html>";
            }
            else
            {
                return File.ReadAllText(msgFile);
            }
        }
        public void SendMail()
        {
            //// Create the Outlook application by using inline initialization.
            Application oApp = new Application();
            // force Outlook to be running using the following 3 statements:
            NameSpace ns = oApp.GetNamespace("MAPI");
            MAPIFolder f = ns.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            System.Threading.Thread.Sleep(5000); // test

            ////Create the new message by using the simplest approach.
            MailItem oMsg = (MailItem)oApp.CreateItem(OlItemType.olMailItem);

            //Add a recipient.
            // TODO: Change the following recipient where appropriate.
            try
            {
                Recipients oRecips = oMsg.Recipients;
                List<string> oTORecip = new List<string>();
                List<string> oCCRecip = new List<string>();
                foreach ( string sendTo in Properties.Settings.Default.sendTo) 
                    oTORecip.Add(sendTo);
                // default To Recipient
                if (oTORecip.Count == 0)
                    oTORecip.Add("douglas.s.elder@boeing.com");
                foreach (string t in oTORecip)
                {
                    Recipient oTORecipt = oRecips.Add(t);
                    oTORecipt.Type = (int)OlMailRecipientType.olTo;
                    oTORecipt.Resolve();
                }

                foreach (string sendCC in Properties.Settings.Default.sendCC)
                    oCCRecip.Add(sendCC);
                foreach (string t in oCCRecip)
                {
                    Recipient oCCRecipt = oRecips.Add(t);
                    oCCRecipt.Type = (int)OlMailRecipientType.olCC;
                    oCCRecipt.Resolve();
                }

                //Set the basic properties.
                oMsg.Subject = "GOLD--MS2--RDMS System Compare Exception Report: " + DateTime.Today.ToString("MM/dd/yyyy");
                oMsg.HTMLBody = getMsg();
                string date = DateTime.Today.ToString("MM-dd-yyyy");

                // If you want to, display the message.
                // oMsg.Display(true);  //modal

                //Send the message.
                oMsg.Save();
                oMsg.Send();
                //Explicitly release objects.
                oTORecip = null;
                //    oCCRecip = null;       
            }
            catch (System.Runtime.InteropServices.COMException e)
            {
                Logger.Error(e,"Mail");
            }

            oMsg = null;
            oApp = null;
        }



    }
}
