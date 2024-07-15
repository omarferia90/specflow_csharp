using AventStack.ExtentReports.Reporter.Configuration;
using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports;
using Microsoft.VisualStudio.TestPlatform.ObjectModel.Client;
using NUnit.Framework;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using log4net;
using log4net.Appender;
using log4net.Config;
using log4net.Core;
using log4net.Layout;

namespace SpecFlow_CSharp.Support
{
    public class ExtentReport
    {
        public static ExtentReports _extentReports;
        public static ExtentTest _featture;
        public static ExtentTest _scenario;
        public static string BaseReportFolder = "";
        public static string FullReportFolder = "";
        public static ILog Logger;

        /// <summary>
        /// Inits report creation and add system information
        /// </summary>
        public static void ExtentReportInit()
        {
            string folderID = $"_{Environment.UserName.ToString()}";
            BaseReportFolder = $@"{TestContext.Parameters["ReportLocation"]}\{DateTime.Now.ToString("MMddyyyy_hhmmss")}{folderID}\";
            FullReportFolder = $@"{BaseReportFolder}{TestContext.Parameters["ProjectName"]}\{TestContext.Parameters["ProjectName"]}.html";
            FolderSetup();
            //var htmlReporter = new ExtentV3HtmlReporter(FullReportFolder);
            var htmlReporter = new ExtentHtmlReporter(FullReportFolder);
            htmlReporter.Config.ReportName = "Automation Status Report";
            htmlReporter.Config.DocumentTitle = "Automation Status Report";
            htmlReporter.Config.Theme = Theme.Dark;
            htmlReporter.Start();
            _extentReports = new ExtentReports();
            _extentReports.AttachReporter(htmlReporter);
            _extentReports.AddSystemInfo("Aplication", TestContext.Parameters["Application"]);
            _extentReports.AddSystemInfo("Browser", TestContext.Parameters["BrowserName"]);
            _extentReports.AddSystemInfo("Env", TestContext.Parameters["TestEnvironment"]);
            _extentReports.AddSystemInfo("OS", "Windows");
            _extentReports.AddSystemInfo("Executed By", folderID.Replace("_", ""));
        }

        /// <summary>
        /// Flush the report.
        /// </summary>
        public static void ExtentReportTearDown()
        {
            _extentReports.Flush();
        }

        /// <summary>
        /// Add an screenshot and returns the location.
        /// </summary>
        /// <param name="driver"></param>
        /// <param name="scenarioContext"></param>
        /// <returns></returns>
        public string addScreenshot(IWebDriver driver, ScenarioContext scenarioContext)
        {
            ITakesScreenshot takesScreenshot = (ITakesScreenshot)driver;
            Screenshot screenshot = takesScreenshot.GetScreenshot();
            string screenshotLocation = Path.Combine($@"{BaseReportFolder}Screenshot", scenarioContext.ScenarioInfo.Title.Replace(".", "") + ".png");
            screenshot.SaveAsFile(screenshotLocation);
            return screenshotLocation;
        }

        /// <summary>
        /// Creates the Report Folder Path.
        /// </summary>
        private static void FolderSetup()
        {
            string[] subFolders = new string[2] { "Screenshot", TestContext.Parameters["ProjectName"] };
            if (!System.IO.Directory.Exists(BaseReportFolder)) { System.IO.Directory.CreateDirectory(BaseReportFolder); }
            foreach (var folder in subFolders)
            { if (!System.IO.Directory.Exists(BaseReportFolder + folder)) { System.IO.Directory.CreateDirectory(BaseReportFolder + folder); } }
        }

        /// <summary>
        /// Add information using log4Net
        /// </summary>
        public static void Log4NetInit()
        {
            PatternLayout patternLayout;
            patternLayout = new PatternLayout();
            patternLayout.ConversionPattern = "%date{ABSOLUTE} [%class] [%level] [%method] - %message%newline";
            patternLayout.ActivateOptions();

            var consoleAppender = new ConsoleAppender()
            {
                Name = "ConsoleAppender",
                Layout = patternLayout,
                Threshold = Level.All
            };

            var fileAppender = new FileAppender()
            {
                Name = "",
                Layout = patternLayout,
                Threshold = Level.All,
                AppendToFile = true,
                File = $@"{BaseReportFolder}{TestContext.Parameters["ProjectName"]}\FileLogger.log"
            };
            fileAppender.ActivateOptions();

            consoleAppender.ActivateOptions();
            BasicConfigurator.Configure(fileAppender);
            Logger = LogManager.GetLogger(typeof(ITestLogger));
        }


    }
}
