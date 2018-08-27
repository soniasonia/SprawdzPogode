namespace SprawdzPogode.Handlers
{
    interface IFetchableHandler: IHandler
    {
        string GetData(string id);
    }
}
