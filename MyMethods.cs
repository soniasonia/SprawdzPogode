using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Xml;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.IO;
using System.Reflection;

namespace SprawdzPogode
{
    public class Control
    {
        private bool error;
        private string description;

        public Control()
        {
            error = false;
            description = "";
        }

        public bool Error
        {
            get; set;
        }
        public string Description
        {
            get; set;
        }

    }
  
    public static class MyMethods
    {
        public static void ReadFromXML(ref string input, ref string output, ref Control con)
        {

         try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load("UserConfig.xml");
                XmlNode node = doc.DocumentElement.SelectSingleNode("//file[@type='input']");
                input = node.InnerText;
                node = doc.DocumentElement.SelectSingleNode("//file[@type='output']");
                output = 
                output = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\" + node.InnerText;

            }
            catch (Exception e)
            {
                con.Error = true;
                con.Description = e.Message;
            }

        }
        public static void ReadFromTXT(ref string input, ref string[] lines, ref Control con)
        {
            try
            {
                lines = System.IO.File.ReadAllLines(input, Encoding.UTF8);
                Console.WriteLine("Input file: " + input);
            }
            catch (Exception e)
            {
                con.Error = true;
                con.Description = e.Message;
            }
        }
        public static void StartExcel(string path, ref Application excel, ref Workbook wb, ref Worksheet ws, ref Control con)
        {
            try
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
            catch (Exception e)
            {
                con.Error = true;
                con.Description = e.Message;
            }
        }
        public static void FinishExcel(ref Application excel)
        {
            excel.Calculation = XlCalculation.xlCalculationAutomatic;
            excel.ScreenUpdating = true;
            excel.EnableEvents = true;
            excel.DisplayAlerts = true;
            

        }
        public static void StartChrome(ref IWebDriver driver, ref Control con)
        {
            try
            {
                driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(60);
                driver.Manage().Window.Maximize();
            }
            catch (Exception e)
            {
                con.Error = true;
                con.Description = e.Message;
            }
        }
        static void preparePage(this IWebDriver driver)
        {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);
        }
        public static IWebElement findElementAfterPreparingPage(this IWebDriver driver, By by, ref Control con)
        {
            preparePage(driver);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
            try
            {
                return driver.FindElement(by);
            }
            catch (OpenQA.Selenium.NoSuchElementException e)
            {
                con.Error = true;
                con.Description = e.Message;
            }
            catch (Exception e)
            {
                con.Error = true;
                con.Description = e.Message;
            }
            return null;
        }
        public static string GetText(IWebElement el)
        {
            if (el != null)
            {
                return el.Text;
            }
            else
            {
                return " - ";
            }
        }
        public static void ExtractDataFromGoogle(ref IWebDriver driver, ref string[] lines, ref Worksheet ws, Control con, LogWriter log)
        {
            IWebElement SearchBox;
            IWebElement Temperatura;
            IWebElement Opady;
            IWebElement Wiatr;

            int LastRow = ws.UsedRange.Rows.Count;
            int row = LastRow +1;
            int counter = 0;
            foreach (string line in lines)
            {
                con.Error = false;
                con.Description = "";
                Console.WriteLine("City: " + line);
                driver.Navigate().GoToUrl("https://www.google.pl/");
                SearchBox = MyMethods.findElementAfterPreparingPage(driver, By.Id("lst-ib"), ref con);
                SearchBox.Click();
                SearchBox.SendKeys(line + " pogoda" + OpenQA.Selenium.Keys.Enter);
                Temperatura = MyMethods.findElementAfterPreparingPage(driver, By.Id("wob_tm"), ref con);
                Opady = MyMethods.findElementAfterPreparingPage(driver, By.Id("wob_pp"), ref con);
                Wiatr = MyMethods.findElementAfterPreparingPage(driver, By.Id("wob_ws"), ref con);

                DateTime now = DateTime.Now;

                ws.Cells[row, 1].Value = String.Format("{0:yyyy/MM/dd HH:mm:ss}", now);
                ws.Cells[row, 2].Value = line;
                ws.Cells[row, 3].Value = MyMethods.GetText(Temperatura);
                ws.Cells[row, 4].Value = MyMethods.GetText(Opady);
                ws.Cells[row, 5].Value = MyMethods.GetText(Wiatr);
                if (con.Error == true)
                {
                    ws.Cells[row, 6].Value = "Fail";
                    ws.Rows[row].Interior.Color = 5296274;
                }
                else
                {
                    ws.Cells[row, 6].Value = "Success";
                }
                log.LogWrite(con, "Extract data for " + line);
                
                row++;
                counter++;
            }
        }
        public static string CheckStatus(Control con, string s)
        {
            if (con.Error == true)
            {
                Console.WriteLine(s + ". Action failed.\n" + con.Description);
                return s + ". Action failed.\n" + con.Description;
            }
            else
            {
                Console.WriteLine(s + ". Action successful.");
                return s + ". Action successful.";
            }

        }
        
}
    public class LogWriter
    {
        public delegate string StatusDel(Control c, string s);
        private string m_exePath = string.Empty;
        public event StatusDel CheckStatus;
        public LogWriter(string logMessage)
        {
            m_exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            using (StreamWriter w = File.AppendText(m_exePath + "\\" + "log.txt"))
            {
                w.WriteLine();
                w.WriteLine(logMessage);
            }
        }
        public void LogWrite(Control con, string action)
        {
            try
            {
                using (StreamWriter w = File.AppendText(m_exePath + "\\" + "log.txt"))
                {
                    string logMessage = action + " " + con.Error + ": " + con.Description;
                    logMessage = CheckStatus(con, action);
                    w.WriteLine(logMessage);
                }
            }
            catch (Exception e)
            {
            }
        }
        public void LogWrite(string logMessage)
        {
            try
            {
                using (StreamWriter w = File.AppendText(m_exePath + "\\" + "log.txt"))
                {
                    w.WriteLine(logMessage);
                }
            }
            catch (Exception e)
            {
            }
        }
    }
}



