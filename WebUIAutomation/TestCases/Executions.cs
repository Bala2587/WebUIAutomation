using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Threading;
using System.Text;
using System.Threading.Tasks;

[assembly: Parallelize(Workers = 100, Scope = ExecutionScope.MethodLevel)]

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

        public ExecuteNow() 
        {
            var dt = GetValues(testcase);
            ToolsQASampleFormPom pom = new ToolsQASampleFormPom(driver);
            if(dt != null)
            {
                foreach(DataColumn col in dt.Columns)
                {
                    if (dt.Rows.ofType<DataRow>().Any(r => r.IsNull(col)))
                        continue;
                }
                else
                {
                    IWebElement element = (IWebElement)ToolsQASampleFormPom.GetType().GetProperty(Col.ColumnName.ToString()).GetValue(pom);
                    
                    if(element.TagName = "input")
                    {
                        element.click();
                        element.SendKeys(dt.Rows[0][col].ToString());
                    }

                    if(element.TagName ="button")
                    {
                        element.click();
                    }
                }
            }
        }

        public DtaTable GetValues(string usecase)
        {
            string filepath = @"c:\\bala\\inflation.xlsx";

            using (var workbook = new XLWorkbook(filepath))
            {
                var worksheet = workbook.Worksheets.First();
                var rows = worksheet
                    .Rows
                    .Skip(1).Where(rows => rows.cell(usecolumn).Value.ToString() == usecase)
                    .Select(rows => GetDictionaryFromRow(worksheet, rows))
                    .ToList();

                var columnHeaders = worksheet
                    .Row(1)
                    .cells(usedCellsOnly: true)
                    .Select(x=>x.Value.ToString())
                    .ToList();
                return ConvertRowstoDataTable(columnHeaders, rows);
            }
        }

        private DataTable ConvertRowstoDataTable(List<string> columns, List<Dictionary<string,string>> rows)
        {
            var dt = new DataTable();

            foreach (var columnName in columns)
            {
                dt.columns.Add(columnName);
            }

            foreach(var row in rows)
            {
                var dataRow = dt.NewRow();
                foreach(var kvp in row)
                {
                    dataRow[kvp.Key] = kvp.Value;
                }

                dt.Rows.Add(dataRow);
            }
            return dt;
        }
    }
}
