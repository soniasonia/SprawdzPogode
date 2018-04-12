using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Windows.Forms;
using SprawdzPogode;

namespace SprawdzPogode
{

    public class Program
    {
        static void Main(string[] args)

        {
            //Ustaw kontrolke
            Control statusControl = new Control();
            LogWriter log = new LogWriter("START");
            log.CheckStatus += MyMethods.CheckStatus;

            //Wczytaj dane z XML-a (sciezki do plikow)
            string input = "";
            string output = "";
            MyMethods.ReadFromXML(ref input, ref output, ref statusControl);
            log.LogWrite(statusControl, "Read from XML");
            if (statusControl.Error == true) return;

            //Wczytaj dane z TXT (lista krajow)
            string[] lines = { };
            MyMethods.ReadFromTXT(ref input, ref lines, ref statusControl);
            log.LogWrite(statusControl, "Read from TXT");
            if (statusControl.Error == true) return;

            //Otworz plik Excel
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = null;
            Worksheet ws = null;
            MyMethods.StartExcel(output, ref excel, ref wb, ref ws, ref statusControl);
            log.LogWrite(statusControl, "Open Excel");
            if (statusControl.Error == true) return;

            //Otworz Chrome
            IWebDriver driver = new ChromeDriver();
            MyMethods.StartChrome(ref driver, ref statusControl);
            log.LogWrite(statusControl, "Open Chrome");
            if (statusControl.Error == true) return;

            //Sprawdz pogode na Google
            MyMethods.ExtractDataFromGoogle(ref driver, ref lines, ref ws,statusControl,log);

            //Zamykanie
            MyMethods.FinishExcel(ref excel);
            wb.Close(true);
            excel.Quit();
            driver.Close();
            Console.ReadKey();
            log.LogWrite("FINISH");
            }

    }
}