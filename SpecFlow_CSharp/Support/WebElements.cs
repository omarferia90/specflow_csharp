using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OpenQA.Selenium.Interactions;

namespace SpecFlow_CSharp.Support
{
    public class WebElements : ExtentReport
    {
        private DefaultWait<IWebDriver> FluentWait;
        private WebDriverWait wait;
        private IWebDriver driver;

        // ********************************************************************************************
        // **                  Standard Functions    -    Most Common Actions
        // ********************************************************************************************

        private object GetFluentWait(IWebElement objWE, string strAction, string strValue = "")
        {
            Actions action = new Actions(driver);
            FluentWait = new DefaultWait<IWebDriver>(driver);
            FluentWait.Timeout = TimeSpan.FromSeconds(5);
            FluentWait.PollingInterval = TimeSpan.FromMilliseconds(250);
            FluentWait.IgnoreExceptionTypes(typeof(WebDriverTimeoutException), typeof(SuccessException));

            switch (strAction.ToUpper())
            {
                case "DISPLAYED":
                    FluentWait.Until(x => objWE.Displayed);
                    break;
                case "SENDKEYS":
                    FluentWait.Until(x => objWE).SendKeys(strValue);
                    break;
                case "CLEAR":
                    FluentWait.Until(x => objWE).Clear();
                    break;
                case "CLICK":
                    FluentWait.Until(x => objWE).Click();
                    break;
                case "DOUBLECLICK":
                    action.DoubleClick(objWE);
                    break;
                case "CUSTOMSENDKEYS":
                    objWE.Click();
                    action.KeyDown(Keys.Control).SendKeys(Keys.Home).Perform();
                    objWE.Clear();
                    Thread.Sleep(TimeSpan.FromMilliseconds(500));
                    objWE.SendKeys(Keys.Delete);
                    objWE.SendKeys(strValue);
                    break;
            }
            return FluentWait;
        }

        public IList<IWebElement> GetWeList(By by)
        {
            try
            {
                ExtentReport.Logger.Debug("Step: Get WebElement List");
                IList<IWebElement> objWE = driver.FindElements(by);
                ExtentReport.Logger.Debug("Step: Get WebElement was found as expected.");
                return objWE;
            }
            catch (Exception objException)
            {
                ExtentReport.Logger.Error("Step: Get WebElement has failed");
                return null;
            }
        }

        public IList<IWebElement> GetWeList(String locator) => GetWeList(By.XPath(locator));

        public IWebElement GetWebElement(By by)
        {
            try
            {
                ExtentReport.Logger.Debug("Step: Get WebElement");
                IWebElement objWE = driver.FindElement(by);
                ExtentReport.Logger.Debug("Step: Get WebElement was found as expected.");

                return objWE;
            }
            catch (Exception objException)
            {
                ExtentReport.Logger.Error("Step: Get WebElement has failed");
                return null;
            }
        }

        public IWebElement GetWebElement(String locator) => GetWebElement(By.XPath(locator));

        public bool WaitPageLoad()
        {
            bool blResult = false;
            try
            {
                ExtentReport.Logger.Debug("Step: Wait page to be loaded.");
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(20));
                IJavaScriptExecutor objJS = (IJavaScriptExecutor)driver;
                wait.Until(wd => objJS.ExecuteScript("return document.readyState").ToString() == "complete");
                blResult = true;
            }
            catch (Exception objException)
            {
            }
            return blResult;
        }

        public bool SendKeys(IWebElement objWE, string strField, string strValue, bool takeScreenshot = false)
        {
            bool blResult = false;
            try
            {
                ExtentReport.Logger.Debug($"Step: Senkkeys for: {strField}.");
                GetFluentWait(objWE, "SendKeys", strValue);
                ExtentReport.Logger.Debug($"Step: Senkkeys for: {strField} was done successfully.");
                blResult = true;
            }
            catch (Exception objException)
            {
                ExtentReport.Logger.Error($"Step: Senkkeys for: {strField} has fail.");
                blResult = false;
            }
            return blResult;
        }

        public bool CustomSendKeys(IWebElement objWE, string strField, string strValue, bool takeScreenshot = false)
        {
            bool blResult = false;
            try
            {
                ExtentReport.Logger.Debug($"Step: CustomSenkkeys for: {strField}.");
                GetFluentWait(objWE, "CustomSendKeys", strValue);
                ExtentReport.Logger.Debug($"Step: CustomSenkkeys for: {strField} was done successfully.");
                blResult = true;
            }
            catch (Exception objException)
            {
                ExtentReport.Logger.Error($"Step: Senkkeys for: {strField} has fail.");
                blResult = false;
            }
            return blResult;
        }

        public bool ClickElement(IWebElement objWE, string strField, bool takeScreenshot = false)
        {
            bool blResult = false;
            try
            {
                ExtentReport.Logger.Debug($"Step: ClickElement for: {strField}.");
                GetFluentWait(objWE, "Click");
                ExtentReport.Logger.Debug($"Step: ClickElement for: {strField} was done successfully.");

                blResult = true;
            }
            catch (Exception objException)
            {
                ExtentReport.Logger.Error($"Step: ClickElement for: {strField} has fail.");
                blResult = false;
            }
            return blResult;
        }

        // ********************************************************************************************
        // **                  Standard Functions    -    Expexted Conditions
        // ********************************************************************************************

        public bool WaitUntilElementPresent(By by, TimeSpan? pTimeToWait = null)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(this.driver, pTimeToWait ?? TimeSpan.FromSeconds(10));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.PresenceOfAllElementsLocatedBy(by));
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool WaitUntilElementPresent(string pLocator, TimeSpan? pTimeToWait = null)
        {
            return WaitUntilElementPresent(By.XPath(pLocator), pTimeToWait);
        }

        public bool ElementDisplayed(By by, TimeSpan? pTimeToWait = null)
        {
            try
            {
                return driver.FindElement(by).Displayed;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool ElementDisplayed(string pLocator, TimeSpan? pTimeToWait = null)
        {
            return ElementDisplayed(By.XPath(pLocator), pTimeToWait);
        }

        public bool WaitUntilElementVisible(By by, TimeSpan? pTimeToWait = null)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(this.driver, pTimeToWait ?? TimeSpan.FromSeconds(10));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(by));
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool WaitUntilElementVisible(string pLocator, TimeSpan? pTimeToWait = null)
        {
            return WaitUntilElementVisible(By.XPath(pLocator), pTimeToWait);
        }

        public bool WaitUntilElementHidden(By by, TimeSpan? pTimeToWait = null)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(this.driver, pTimeToWait ?? TimeSpan.FromSeconds(40));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.InvisibilityOfElementLocated(by));
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool WaitUntilElementHidden(string pLocator, TimeSpan? pTimeToWait = null)
        {
            return WaitUntilElementHidden(By.XPath(pLocator), pTimeToWait);
        }

        public bool WaitUntilElementClickable(By by, TimeSpan? pTimeToWait = null)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(this.driver, pTimeToWait ?? TimeSpan.FromSeconds(40));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(by));
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool WaitUntilElementClickable(string pLocator, TimeSpan? pTimeToWait = null)
        {
            return WaitUntilElementClickable(By.XPath(pLocator), pTimeToWait);
        }

        public bool WaitUntilElementSelected(By by, TimeSpan? pTimeToWait = null)
        {
            try
            {
                WebDriverWait wait = new WebDriverWait(this.driver, pTimeToWait ?? TimeSpan.FromSeconds(25));
                wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeSelected(by));
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        public bool WaitUntilElementSelected(string pLocator, TimeSpan? pTimeToWait = null)
        {
            return WaitUntilElementSelected(By.XPath(pLocator), pTimeToWait);
        }

        public string GetHexColor(IWebElement objWE, string attribute = "background-color")
        {
            string hexval;

            try
            {
                String[] numbers;
                String color = objWE.GetCssValue(attribute);
                numbers = color.Contains("rgba") ? color.Replace("rgba(", "").Replace(")", "").Split(',') : color.Replace("rgb(", "").Replace(")", "").Split(',');
                int r = Convert.ToInt16(numbers[0].Trim());
                int g = Convert.ToInt16(numbers[1].Trim());
                int b = Convert.ToInt16(numbers[2].Trim());
                hexval = $"#{r.ToString("X2")}{g.ToString("X2")}{b.ToString("X2")}";
            }
            catch (Exception objException)
            {
                hexval = "";
            }
            return hexval;
        }


    }
}
