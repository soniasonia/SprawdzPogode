using SprawdzPogode.Handlers;

namespace SprawdzPogode.Extractors
{
    interface IExtractor
    {
        void ExtractData(IOutputHandler ex, IFetchableHandler ch);
    }
}
