using System.Xml;
using SprawdzPogode.Readers;
using SprawdzPogode.Handlers;
using OpenQA.Selenium.Chrome;
using Microsoft.Office.Interop.Excel;
using SprawdzPogode.Extractors;

namespace SprawdzPogode
{
    public class Program
    {
        const string XMLConfigPath = "UserConfig.xml";

        static void Main(string[] args)
        {
            XMLReader config = new XMLReader(XMLConfigPath, new XmlDocument());
            string[] paths = config.Read();
            TXTReader txt = new TXTReader(paths[0]);
            ExcelHandler ex = new ExcelHandler(paths[1], new Application());
            ChromeHandler ch = new ChromeHandler(new ChromeDriver());

            Extractor ext = new Extractor(txt.Read());
            ext.ExtractData(ex, ch);
        }
    }
}