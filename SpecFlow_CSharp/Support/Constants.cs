using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpecFlow_CSharp.Support
{
    public class Constants
    {
        //Add Configuration Report
        #region Report
        public static string ReportLocation = @"C:\\AutomationProject\\Report";
        public static string ReportName = "Automated Report";
        public static string ProjectName = "";
        public static string TestEnvironment = "";
        public static string TestBrowser = "";
        #endregion

        //Add Configuration DataDriver
        #region DataDriver
        public static string DataDriverPath = @"C:\AutomationProjects\Data\";
        public static string DataDriverName = "AutomationDriver.xlsx";
        public static string DataDriverSheet = "TestCases";
        #endregion

        //Add Environments URL
        public static string UrlApp = "https://open.spotify.com/intl-es";
        public static string UrlExplorer = "https://apps.powerapps.com/play/e/default-d7063cf2-afd6-43fc-aa12-1667d97f0885/a/1116cea8-fdf5-4d29-9961-80f434f60a37?tenantId=d7063cf2-afd6-43fc-aa12-1667d97f0885&source=sharebutton&sourcetime=1721069078083";

    }
}
