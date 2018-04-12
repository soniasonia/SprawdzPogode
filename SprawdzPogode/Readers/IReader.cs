namespace SprawdzPogode.Readers
{
    interface IReader
    {
        string Path { get; set; }

        string[] Read();
    }
}
