using System;
using System.Text;
using System.IO;
using System.Reflection;
using System.Xml;

namespace SprawdzPogode
{
    public interface IReader
    {
        void Read();
    }

    public class XMLConfig : IReader
    {
        private string input;
        private string output;
        public string Input
        {
            get
            {
                return input;
            }
        }
        public string Output
        {
            get
            {
                return output;
            }
        }

        public void Read()
        {
            XmlDocument doc = new XmlDocument();
            doc.Load("UserConfig.xml");
            XmlNode node = doc.DocumentElement.SelectSingleNode("//file[@type='input']");
            this.input = node.InnerText;
            node = doc.DocumentElement.SelectSingleNode("//file[@type='output']");
            this.output = node.InnerText;
        }
    }

    public class MyTXTReader : IReader
    {
        private string[] lines;
        private string path;

        public string[] Lines
        {
            get
            {
                return lines;
            }
        }

        public MyTXTReader(string path)
        {
            this.path = path;
        }

        public void Read()
        {
            lines = System.IO.File.ReadAllLines(this.path, Encoding.UTF8);
            Console.WriteLine("Input file: " + this.path);
        }
    }
}