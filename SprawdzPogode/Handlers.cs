using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;

namespace SprawdzPogode
{
    public class ExcelHandler
    {
        public ExcelHandler(string path)
        {
            this.path = path;
        }
        private Application excel;
        private Workbook wb;
        private Worksheet ws;
        private string path;

        public Worksheet Ws
        {
            get
            {
                return ws;
            }
        }

        public void Start()
        {
            Console.WriteLine("Output file: " + path);
            excel = new Application();
            excel.Visible = true;
            wb = excel.Workbooks.Open(path);
            ws = excel.ActiveSheet as Worksheet;
            excel.Calculation = XlCalculation.xlCalculationManual;
            excel.ScreenUpdating = false;
            excel.EnableEvents = false;
            excel.DisplayAlerts = false;
        }
        public void Finish()
        {
            excel.Calculation = XlCalculation.xlCalculationAutomatic;
            excel.ScreenUpdating = true;
            excel.EnableEvents = true;
            excel.DisplayAlerts = true;
            wb.Save();
            excel.Quit();
        }
    }
    public class ChromeHandler
    {
        private IWebDriver driver;
        private IWebElement el;

        public void Start()
        {
        driver = new ChromeDriver();
        driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(60);
        driver.Manage().Window.Maximize();
        }
        public IWebElement findElement(By by)
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            return driver.FindElement(by);
        }
        public void Search(string input)
        {
            driver.Navigate().GoToUrl("https://www.google.pl/");
            IWebElement SearchBox = findElement(By.Id("lst-ib"));
            SearchBox.Click();
            SearchBox.SendKeys(input + " pogoda" + OpenQA.Selenium.Keys.Enter);
        }
        public string GetData(By by)
        {
            try
            {
                el = findElement(by);
            }
            catch (OpenQA.Selenium.NoSuchElementException e)
            {
                el = null;
            }
            if (el != null)
            {
                return el.Text;
            }
            else
            {
                return "[not found]";
            }
        }
        public void Finish()
        {
            driver.Close();
        }
    }
}





