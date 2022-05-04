using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;

namespace WebUIAutomation
{
    [TestClass]
    public class BaseTest
    {
        public BaseTest()
        {

        }
        public IWebDriver driver;

        [TestInitialize]
        public void initializeDriver()
        {
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("start-maximized");
            driver = new ChromeDriver(options);
        }

        [TestCleanup]
        public void TestCleanup()
        {
            driver.Manage().Cookies.DeleteAllCookies();
            driver.Dispose();
        }
    }
}
