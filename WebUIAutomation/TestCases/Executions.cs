using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Text;
using System.Threading.Tasks;


namespace WebUIAutomation
{
    [TestClass]
    public class Executions: BaseTest
    {
        [TestMethod]
        public void FillFormwithHardcodedValues()
        {
            WebDriverWait Wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));

            driver.Navigate().GoToUrl("https://demoqa.com/automation-practice-form");
            ToolsQASampleFormPom pom = new ToolsQASampleFormPom(driver);
            Wait.Until(ExpectedConditions.UrlContains("https://demoqa.com/automation-practice-form"));

            pom.FirstName.SendKeys("ABC");
            pom.LastName.SendKeys("XYZ");
            pom.Email.SendKeys("abc@gmail.com");
            pom.Gender("Male").Click();
            pom.Mobile.SendKeys("9911223344");
            Thread.Sleep(6000); 
            ((OpenQA.Selenium.ITakesScreenshot)driver).GetScreenshot().SaveAsFile("FillFormwithHardcodedValues.png"); // image will be stored in bin folder
        }
    }
}
