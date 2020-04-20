using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Threading;
using ClosedXML;


namespace SpeedyTestProject1
{
    public class Tests
    {

        IWebDriver driver = new ChromeDriver();

        
        [SetUp]
        public void Setup()
        {
            driver.Navigate().GoToUrl("https://services.speedy.bg/calculate/");
            
            Console.WriteLine("Opened Speedy.bg/calculate/");

            

        }

        [Test]
        public void Test1()
        {

         
            //string sAddress = xRange.Cells[1,1].ToString();

            IWebElement senderAddress = driver.FindElement(By.Id("sndrSiteName"));
            senderAddress.SendKeys("СЛИВЕН");
            System.Threading.Thread.Sleep(1500);
            senderAddress.SendKeys(Keys.Enter);

            System.Threading.Thread.Sleep(1500);

            IWebElement dropToOffice = driver.FindElement(By.Id("dropoffOffice"));
            dropToOffice.Click();

            System.Threading.Thread.Sleep(1500);

            IWebElement recipientAddres = driver.FindElement(By.Id("rcptSiteName"));
            recipientAddres.SendKeys("СОФИЯ");
            System.Threading.Thread.Sleep(1500);
            recipientAddres.SendKeys(Keys.Enter);

            System.Threading.Thread.Sleep(1500);

            IWebElement pickFromOffice = driver.FindElement(By.Id("pickupOffice"));
            pickFromOffice.Click();

            System.Threading.Thread.Sleep(1500);

            IWebElement packageType = driver.FindElement(By.Id("packageType"));
            packageType.Click();

            System.Threading.Thread.Sleep(1500);

            IWebElement isDocuments = driver.FindElement(By.Id("documents"));
            isDocuments.Click();

            System.Threading.Thread.Sleep(1500);

            IWebElement calculateButton = driver.FindElement(By.CssSelector(".inputButtonCalc"));
            calculateButton.Click();

            System.Threading.Thread.Sleep(1500);

        }

        [TearDown]
        public void CleanUp()
        {
            driver.Close();
            Console.WriteLine("Driver closed");
        }



    }
}