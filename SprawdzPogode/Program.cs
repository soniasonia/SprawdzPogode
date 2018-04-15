using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;

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
           /* Control statusControl = new Control();
            LogWriter log = new LogWriter("START");
            log.CheckStatus += MyMethods.CheckStatus;*/


                XMLConfig config = new XMLConfig(); //wczytaj sciezki z XML
                config.Read();
                MyTXTReader txt = new MyTXTReader(config.Input); //wczytaj miasta z TXT
                txt.Read();
                ExcelHandler ex = new ExcelHandler(config.Output);
                ex.Start();
                ChromeHandler ch = new ChromeHandler();
                ch.Start();
                Extractor ext = new Extractor(txt.Lines);
                ext.ExtractData(ex,ch);
                ex.Finish();
                ch.Finish();

            /*
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
                        */
        }
    }
}