using System.Text;

namespace SprawdzPogode.Readers
{
    public class TXTReader : IReader
    {
        public string Path { get; set; }

        public TXTReader(string path)
        {
            Path = path;
        }

        public string[] Read()
        {
            return System.IO.File.ReadAllLines(Path, Encoding.UTF8);
        }
    }
}