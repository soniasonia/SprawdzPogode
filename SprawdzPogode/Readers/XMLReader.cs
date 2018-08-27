using System.Xml;

namespace SprawdzPogode.Readers
{
    public class XMLReader : IReader
    {
        public string Path { get; set; }
        public XmlDocument XmlDataDocument { get; set; }

        public XMLReader(string path, XmlDocument XmlDocument)
        {
            Path = path;
            XmlDataDocument = XmlDocument;
        }

        public string[] Read()
        {
            XmlDataDocument.Load(Path);

            return new string[] {
                XmlDataDocument.DocumentElement.SelectSingleNode("//file[@type='input']").InnerText,
                XmlDataDocument.DocumentElement.SelectSingleNode("//file[@type='output']").InnerText
            };
        }
    }
}
