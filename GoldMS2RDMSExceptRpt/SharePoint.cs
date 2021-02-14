using System;

using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using System.Windows.Forms;
using System.Diagnostics;
using System.Net;
using System.Security.Principal;
using System.Configuration;




namespace GoldMS2RDMSExceptRpt
{
    [Serializable]
    class BadExceptRptFile : Exception
    {
        public BadExceptRptFile(string message)
            : base(message)
        {
        }
    }
    class SharePoint
    {
        public void upload(String fileName, String sharePointUNC)
        {
            // for the copy to work we need to connect to the sharePointUNC
            String sharePointDrive = "Q";
            if (Properties.Settings.Default.sharePointDrive != null)
                sharePointDrive = Properties.Settings.Default.sharePointDrive;
            DriveSettings.MapNetworkDrive(sharePointDrive, sharePointUNC);

            try
            {
                String basename = Path.GetFileName(fileName);
                String destFile = System.IO.Path.Combine(sharePointUNC, basename);
                if (System.IO.File.Exists(destFile))
                {
                    DateTime lastModified = System.IO.File.GetLastWriteTime(destFile);
                    DateTime mountain = TimeZoneInfo.ConvertTimeBySystemTimeZoneId
                     (DateTime.Now, "Mountain Standard Time");

                    if (lastModified.AddDays(7).Date != mountain.Date )
                    {
                        throw new BadExceptRptFile(
                            string.Format("The file {0} was already created on {1}", 
                            basename, lastModified.ToString("MM/dd/yy")));
                    }
                }
                System.IO.File.Copy(fileName, destFile, true);
                Logger.Info(fileName + " copied to " + destFile,"SharePoint");
            } 
            catch (BadExceptRptFile e)
            {
                throw e;
            }
            catch (Exception e)
            {
                Logger.Error(e, "SharePoint");
                Application.Exit();
            }
            // the mapped drive is no longer needed
            if (DriveSettings.IsDriveMapped(sharePointDrive))
                DriveSettings.DisconnectNetworkDrive(sharePointDrive, true);
        }

    }
}
