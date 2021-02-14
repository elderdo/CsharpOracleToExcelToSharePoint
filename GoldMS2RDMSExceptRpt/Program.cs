using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Deployment.Application;

namespace GoldMS2RDMSExceptRpt
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();


            Application.SetCompatibleTextRenderingDefault(false);
            if (args.Length > 0)
            {
                Application.Run(new propertyEditor());
                Application.Exit();
                return;
            }
            DateTime mountain = TimeZoneInfo.ConvertTimeBySystemTimeZoneId
                (DateTime.Now, "Mountain Standard Time");
            DayOfWeek today = mountain.DayOfWeek;

            if (today == Properties.Settings.Default.runDay
                && mountain.Hour >= Properties.Settings.Default.startTime
                && mountain.Hour <= Properties.Settings.Default.endTime)
            {
                Application.Run(new Login());
            }
            else
            {
                string version = Properties.Settings.Default.appName;
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    version = string.Format("{0} - v{1}",version, ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString(4));
                }
                Logger.Error("This application should only run on "
                    + Properties.Settings.Default.runDay
                    + " between "
                    + Properties.Settings.Default.startTime
                    + " and "
                    + Properties.Settings.Default.endTime, version);
            }

        }
    }
}
