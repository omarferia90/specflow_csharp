using AventStack.ExtentReports;
using AventStack.ExtentReports.Gherkin.Model;
using BoDi;
using OpenQA.Selenium;
using SpecFlow_CSharp.Drivers;
using SpecFlow_CSharp.Support;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpecFlow_CSharp.Hooks
{
    [Binding]
    public sealed class Hooks : ExtentReport
    {

        private readonly IObjectContainer _container;
        public Hooks(IObjectContainer container)
        {
            _container = container;
        }


        [BeforeTestRun]
        public static void BeforeTestRun()
        {
            ExtentReportInit();
            Log4NetInit();
        }


        [AfterTestRun]
        public static void AfterTestRun()
        { ExtentReportTearDown(); }


        [BeforeFeature]
        public static void BeforeFeature(FeatureContext featureContext)
        { _featture = _extentReports.CreateTest<Feature>(featureContext.FeatureInfo.Title); }


        [BeforeScenario(Order = 1)]
        public void FirstBeforeScenario(ScenarioContext scenarioContext)
        {
            //IWebDriver driver = new ChromeDriver();
            WebDriverFactory.CreateDriverManager(BrowserType.Chrome);
            IWebDriver driver = WebDriverFactory.CreateDriver(BrowserType.Chrome);
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(20);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
            driver.Manage().Window.Maximize();
            driver.Manage().Cookies.DeleteAllCookies();
            _container.RegisterInstanceAs<IWebDriver>(driver);
            _scenario = _featture.CreateNode<Scenario>(scenarioContext.ScenarioInfo.Title);
            ExtentReport.Logger.Debug($"***************************************************************");
            ExtentReport.Logger.Debug($"Step: Scenario >>> {scenarioContext.ScenarioInfo.Title} starts.");
            ExtentReport.Logger.Debug($"***************************************************************");
        }


        [AfterScenario]
        public void AfterScenario(ScenarioContext scenarioContext)
        {
            var driver = _container.Resolve<IWebDriver>();
            if (driver != null) { driver.Quit(); }
            ExtentReport.Logger.Debug($"***************************************************************");
            ExtentReport.Logger.Debug($"Step: Scenario >>> {scenarioContext.ScenarioInfo.Title} ends.");
            ExtentReport.Logger.Debug($"***************************************************************");
        }


        [AfterStep]
        public void AfterStep(ScenarioContext scenarioContext)
        {
            string stepType = scenarioContext.StepContext.StepInfo.StepDefinitionType.ToString();
            string stepName = scenarioContext.StepContext.StepInfo.Text;
            var driver = _container.Resolve<IWebDriver>();

            //When an exception occurs but scenarioContext.TestError is null
            if (scenarioContext.ScenarioExecutionStatus != ScenarioExecutionStatus.OK)
            {
                addScreenshot(driver, scenarioContext);
                var exceptionMsg = $"An exception occurred and force the step to fail.";
                if (stepType == "Given") { _scenario.CreateNode<Given>(stepName).Fail(exceptionMsg, MediaEntityBuilder.CreateScreenCaptureFromPath(addScreenshot(driver, scenarioContext)).Build()); }
                else if (stepType == "When") { _scenario.CreateNode<When>(stepName).Fail(exceptionMsg, MediaEntityBuilder.CreateScreenCaptureFromPath(addScreenshot(driver, scenarioContext)).Build()); }
                else if (stepType == "Then") { _scenario.CreateNode<Then>(stepName).Fail(exceptionMsg, MediaEntityBuilder.CreateScreenCaptureFromPath(addScreenshot(driver, scenarioContext)).Build()); }
            }

            //When Scenario Fails
            if (scenarioContext.TestError != null)
            {
                addScreenshot(driver, scenarioContext);
                if (stepType == "Given") { _scenario.CreateNode<Given>(stepName).Fail(scenarioContext.TestError.Message, MediaEntityBuilder.CreateScreenCaptureFromPath(addScreenshot(driver, scenarioContext)).Build()); }
                else if (stepType == "When") { _scenario.CreateNode<When>(stepName).Fail(scenarioContext.TestError.Message, MediaEntityBuilder.CreateScreenCaptureFromPath(addScreenshot(driver, scenarioContext)).Build()); }
                else if (stepType == "Then") { _scenario.CreateNode<Then>(stepName).Fail(scenarioContext.TestError.Message, MediaEntityBuilder.CreateScreenCaptureFromPath(addScreenshot(driver, scenarioContext)).Build()); }
            }

            //When Scenario Passed
            if (scenarioContext.TestError == null)
            {
                if (stepType == "Given") { _scenario.CreateNode<Given>(stepName); }
                else if (stepType == "When") { _scenario.CreateNode<When>(stepName); }
                else if (stepType == "Then") { _scenario.CreateNode<Then>(stepName); }
            }
        }

    }
}
