namespace SprawdzPogode.Handlers
{
    interface IHandler
    {
        void Start();
        void Finish();
        void Handle(string[] values);
    }
}