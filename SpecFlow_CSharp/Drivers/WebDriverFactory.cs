using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;

namespace SpecFlow_CSharp.Drivers
{
    /// <summary>
    /// Enum to select specific browser
    /// </summary>
    public enum BrowserType
    {
        Chrome,
        Firefox,
        Edge
    }

    /// <summary>
    /// Creates a web browser intance 
    /// </summary>
    public class WebDriverFactory
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="browserType"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static IWebDriver CreateDriver(BrowserType browserType)
        {
            switch (browserType)
            {
                case BrowserType.Chrome:
                    return new ChromeDriver();
                case BrowserType.Firefox:
                    return new FirefoxDriver();
                case BrowserType.Edge:
                    return new EdgeDriver();
                default:
                    throw new ArgumentException("Invalid browser type");
            }
        }

        /// <summary>
        /// Creates a driver manager instance
        /// </summary>
        /// <param name="browserType"></param>
        /// <returns></returns>
        /// <exception cref="ArgumentException"></exception>
        public static string CreateDriverManager(BrowserType browserType)
        {
            DriverManager driverManager = new DriverManager();
            switch (browserType)
            {
                case BrowserType.Chrome:
                    ChromeOptions chromeOptions = new ChromeOptions();
                    chromeOptions.AddArgument("no-sandbox");
                    return driverManager.SetUpDriver(new ChromeConfig());
                case BrowserType.Firefox:
                    FirefoxOptions firefoxOptions = new FirefoxOptions();
                    firefoxOptions.AddArgument("no-sandbox");
                    return driverManager.SetUpDriver(new FirefoxConfig());
                case BrowserType.Edge:
                    EdgeOptions edgeOptions = new EdgeOptions();
                    edgeOptions.AddArgument("no-sandbox");
                    return driverManager.SetUpDriver(new EdgeConfig());
                default:
                    throw new ArgumentException("Invalid browser type");
            }
        }


    }
}
