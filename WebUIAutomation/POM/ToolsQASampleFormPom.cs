using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.PageObjects;
using SeleniumExtras.WaitHelpers;

namespace WebUIAutomation
{
   
    public class ToolsQASampleFormPom
    {
        IWebDriver driver;
        WebDriverWait Wait;

        public ToolsQASampleFormPom(IWebDriver driver)
        {
            this.driver = driver;
            this.Wait = new WebDriverWait(driver, TimeSpan.FromSeconds(60));
            Wait.Until(ExpectedConditions.UrlContains("https://demoqa.com/automation-practice-form"));
        }

        public IWebElement FirstName => driver.FindElement(By.CssSelector("input[id*= 'firstName']"));
        public IWebElement LastName => driver.FindElement(By.XPath("//*[@id='lastName']"));
        public IWebElement Email => driver.FindElement(By.CssSelector("input[id*= 'userEmail']"));
        public IWebElement Gender(string gender) => driver.FindElement(By.XPath("//*[text() ='"+gender+"']"));
        public IWebElement Mobile => driver.FindElement(By.CssSelector("input[id*= 'userNumber']"));
        public IWebElement DOB => driver.FindElement(By.CssSelector("input[id*= 'dateOfBirthInput']"));
        public IWebElement Subjects => driver.FindElement(By.CssSelector("input[id*= 'subjectsInput']"));
        public IWebElement Hobbies(string hobby) => driver.FindElement(By.XPath("//*[text() ='" + hobby + "']"));
        public IWebElement state => driver.FindElement(By.CssSelector("div[id*= 'state']"));
        public IWebElement city => driver.FindElement(By.CssSelector("div[id*= 'city']"));




    }
}
