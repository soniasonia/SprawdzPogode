using SprawdzPogode.Exceptions;
using System;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;

namespace SprawdzPogode.Handlers
{
    public class ChromeHandler : IFetchableHandler
    {
        const string GoogleUrl = "https://www.google.pl/";
        public ChromeDriver Driver { set; get; }

        public ChromeHandler(ChromeDriver driver)
        {
            Driver = driver;
        }

        public void Start()
        {
            Driver.Manage().Timeouts().PageLoad = TimeSpan.FromSeconds(60);
            Driver.Manage().Window.Maximize();
            Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }

        public void Handle(string[] values)
        {
            Driver.Navigate().GoToUrl(GoogleUrl);
            IWebElement SearchBox = Driver.FindElement(By.Id("lst-ib"));
            SearchBox.Click();
            SearchBox.SendKeys(values[0] + Keys.Enter);
        }

        public string GetData(string id)
        {
            try
            {
                return Driver.FindElement(By.Id(id)).Text;
            }
            catch (OpenQA.Selenium.NoSuchElementException e)
            {
                throw new DataNotFoundException();
            }
        }

        public void Finish()
        {
            Driver.Close();
        }
    }
}





