using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Threading;
using ClosedXML;
using ClosedXML.Excel;
using System.Collections.Generic;
using System.IO;

namespace SpeedyTestProject1
{
    public class Tests
    {

     /*   private static IEnumerable<NUnitTestProject1.All> Data()
        {
            return new List<NUnitTestProject1.All>
            {
                new All{Id = 1, Cases= 1, Date = DateTime.Now, Deaths= 0, Recovered=0},
                new All{Id = 2, Cases= 2, Date = DateTime.Now, Deaths= 1, Recovered=1},
                new All{Id = 3, Cases= 3, Date = DateTime.Now, Deaths= 1, Recovered=2}
            };
        }*/


        IWebDriver driver = new ChromeDriver();

        
        [SetUp]
        public void Setup()
        {
            driver.Navigate().GoToUrl("https://services.speedy.bg/calculate/");
            
            Console.WriteLine("Opened Speedy.bg/calculate/");
            
        }

        
        [Test]
        public void Test1_InputCorrectDataForDocumentsPackageAndReturnPrice()
        {

            var fileName = Path.Combine(Environment.CurrentDirectory, "Data\\speedy.xlsx");
            var workbook = new XLWorkbook(fileName);
            var ws1 = workbook.Worksheet(1);
            int iRow = 1;
            while (!ws1.Cell(iRow, 1).IsEmpty())
            {
                var row = "";
                int iColumn = 1;
                while (!ws1.Cell(iRow, iColumn).IsEmpty())
                {
                    row = row + ws1.Cell(iRow, iColumn).Value.ToString() + ",";
                    //string tempValue = ws1.Cell(iRow, iColumn).Value.ToString();
                    //Console.OutputEncoding = Encoding.UTF8;
                    //Console.WriteLine(ws1.Cell(iRow, iColumn).Value);
                    iColumn++;
                }
                //Console.WriteLine(row);
                iRow++;
            }

            string sAddress = ws1.Cell(2, 1).Value.ToString();
            string rAddress = ws1.Cell(2, 2).Value.ToString();

            IWebElement senderAddress = driver.FindElement(By.Id("sndrSiteName"));
            senderAddress.SendKeys(sAddress);
            System.Threading.Thread.Sleep(1500);
            senderAddress.SendKeys(Keys.Enter);

            System.Threading.Thread.Sleep(1500);

            IWebElement dropToOffice = driver.FindElement(By.Id("dropoffOffice"));
            dropToOffice.Click();

            System.Threading.Thread.Sleep(1500);

            IWebElement recipientAddres = driver.FindElement(By.Id("rcptSiteName"));
            recipientAddres.SendKeys(rAddress);
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

            System.Threading.Thread.Sleep(5000);

            IWebElement priceWithVAT = driver.FindElement(By.CssSelector(".tdRight"));
            string price = priceWithVAT.Text;

            ws1.Cell(2, 10).Value = price;

            workbook.SaveAs(fileName);


        }

        [Test]
        public void Test2_InputCorrectDataForPalletAndSelectTypeBasedOnInput()
        {

            var fileName = Path.Combine(Environment.CurrentDirectory, "Data\\speedy.xlsx");
            var workbook = new XLWorkbook(fileName);
            var ws1 = workbook.Worksheet(1);
            int iRow = 1;
            while (!ws1.Cell(iRow, 1).IsEmpty())
            {
                var row = "";
                int iColumn = 1;
                while (!ws1.Cell(iRow, iColumn).IsEmpty())
                {
                    row = row + ws1.Cell(iRow, iColumn).Value.ToString() + ",";
                    //string tempValue = ws1.Cell(iRow, iColumn).Value.ToString();
                    //Console.OutputEncoding = Encoding.UTF8;
                    //Console.WriteLine(ws1.Cell(iRow, iColumn).Value);
                    iColumn++;
                }
                //Console.WriteLine(row);
                iRow++;
            }

            string sAddress = ws1.Cell(4, 1).Value.ToString();
            string rAddress = ws1.Cell(4, 2).Value.ToString();

            IWebElement senderAddress = driver.FindElement(By.Id("sndrSiteName"));
            senderAddress.SendKeys(sAddress);
            System.Threading.Thread.Sleep(1500);
            senderAddress.SendKeys(Keys.Enter);

            System.Threading.Thread.Sleep(1500);

            IWebElement dropToOffice = driver.FindElement(By.Id("dropoffOffice"));
            dropToOffice.Click();

            System.Threading.Thread.Sleep(1500);

            IWebElement recipientAddres = driver.FindElement(By.Id("rcptSiteName"));
            recipientAddres.SendKeys(rAddress);
            System.Threading.Thread.Sleep(1500);
            recipientAddres.SendKeys(Keys.Enter);

            System.Threading.Thread.Sleep(1500);

            IWebElement pickFromOffice = driver.FindElement(By.Id("pickupOffice"));
            pickFromOffice.Click();

            System.Threading.Thread.Sleep(1500);

            IWebElement packageType = driver.FindElement(By.Id("packageType"));
            packageType.Click();

            System.Threading.Thread.Sleep(1500);

            //IWebElement isDocuments = driver.FindElement(By.Id("documents"));
            //isDocuments.Click();

            System.Threading.Thread.Sleep(1500);

            string d1 = ws1.Cell(4, 3).Value.ToString();
            string d2 = ws1.Cell(4, 4).Value.ToString();
            string d3 = ws1.Cell(4, 5).Value.ToString();
            string tWeight = ws1.Cell(4, 6).Value.ToString();

            IWebElement dimension1 = driver.FindElement(By.Id("D1"));
            dimension1.SendKeys(d1);

            IWebElement dimension2 = driver.FindElement(By.Id("D2"));
            dimension2.SendKeys(d2);

            IWebElement dimension3 = driver.FindElement(By.Id("D3"));
            dimension3.SendKeys(d3);

            IWebElement totalWeight = driver.FindElement(By.Id("totalWeight"));
            totalWeight.SendKeys(tWeight);

            IWebElement payment = driver.FindElement(By.Id("cod"));
            payment.SendKeys("100");


            IWebElement calculateButton = driver.FindElement(By.CssSelector(".inputButtonCalc"));
            calculateButton.Click();

            System.Threading.Thread.Sleep(5000);
                        
            IWebElement priceWithVAT = driver.FindElement(By.CssSelector(".tdRight"));
            string price = priceWithVAT.Text;

            ws1.Cell(4, 10).Value = price;

            workbook.SaveAs(fileName);

            Console.WriteLine(price);


        }

        [Test]
        public void Test3_InputCorrectDataForPalletAndSelectTypeBasedOnInputThenReturnPrice()
        {
            var fileName = Path.Combine(Environment.CurrentDirectory, "Data\\speedy.xlsx");
            //var fileName = Path.Combine(@"C:\Users\smasl\source\repos\ReadFromExcel\ReadFromExcel\Data", "\\speedy.xlsx");
            //C:\Users\smasl\source\repos\ReadFromExcel\ReadFromExcel\Data
            var workbook = new XLWorkbook(fileName);
            var ws1 = workbook.Worksheet(1);
            int iRow = 1;
            while (!ws1.Cell(iRow, 1).IsEmpty())
            {
                var row = "";
                int iColumn = 1;
                while (!ws1.Cell(iRow, iColumn).IsEmpty())
                {
                    row = row + ws1.Cell(iRow, iColumn).Value.ToString() + ",";
                    //string tempValue = ws1.Cell(iRow, iColumn).Value.ToString();
                    //Console.OutputEncoding = Encoding.UTF8;
                    //Console.WriteLine(ws1.Cell(iRow, iColumn).Value);
                    iColumn++;
                }
                //Console.WriteLine(row);
                iRow++;
            }

            string sAddress = ws1.Cell(3, 1).Value.ToString();
            string rAddress = ws1.Cell(3, 2).Value.ToString();

            IWebElement senderAddress = driver.FindElement(By.Id("sndrSiteName"));
            senderAddress.SendKeys(sAddress);
            System.Threading.Thread.Sleep(1500);
            senderAddress.SendKeys(Keys.Enter);

            System.Threading.Thread.Sleep(1500);

            IWebElement dropToOffice = driver.FindElement(By.Id("dropoffOffice"));
            dropToOffice.Click();

            System.Threading.Thread.Sleep(1500);

            IWebElement recipientAddres = driver.FindElement(By.Id("rcptSiteName"));
            recipientAddres.SendKeys(rAddress);
            System.Threading.Thread.Sleep(1500);
            recipientAddres.SendKeys(Keys.Enter);

            System.Threading.Thread.Sleep(1500);

            IWebElement pickFromOffice = driver.FindElement(By.Id("pickupOffice"));
            pickFromOffice.Click();

            System.Threading.Thread.Sleep(1500);

            IWebElement palletType = driver.FindElement(By.Id("palletType"));
            palletType.Click();

            string sPalletWeight = ws1.Cell(3, 7).Value.ToString();

            IWebElement palletWeight = driver.FindElement(By.Id("baseWeight"));
            palletWeight.SendKeys(sPalletWeight);

            string sPalletType = ws1.Cell(3, 8).Value.ToString();

            if (sPalletType == "150") {

                new SelectElement(driver.FindElement(By.Id("baseName"))).SelectByText("150 - Mini Pallet (EUR6) (80 x 60)");

            } if (sPalletType == "250")
            {

                new SelectElement(driver.FindElement(By.Id("baseName"))).SelectByText("250 - Euro Pallet (80 x 120)");

            } if (sPalletType == "350")
                {

                    new SelectElement(driver.FindElement(By.Id("baseName"))).SelectByText("350 - Industrial Pallet 100 (EUR3) (100 x 120)");

            } if (sPalletType == "450")
                {

                    new SelectElement(driver.FindElement(By.Id("baseName"))).SelectByText("450 - Industrial Pallet 120 (120 x 120)");

                }


            System.Threading.Thread.Sleep(1500);

            string sPalletHeight = ws1.Cell(3, 9).Value.ToString();

            IWebElement palletHeight = driver.FindElement(By.Id("baseHeight"));
            palletHeight.SendKeys(sPalletHeight);

            System.Threading.Thread.Sleep(1500);

            IWebElement palletCount = driver.FindElement(By.Id("baseCount"));
            palletCount.SendKeys("1");

            IWebElement addPallet = driver.FindElement(By.CssSelector(".inputPalletBaseAddButton"));
            addPallet.Click();


            System.Threading.Thread.Sleep(1500);

            IWebElement calculateButton = driver.FindElement(By.CssSelector(".inputButtonCalc"));
            calculateButton.Click();

            System.Threading.Thread.Sleep(1500);

            IWebElement priceWithVAT = driver.FindElement(By.CssSelector("#divStatus > table > tbody > tr.trRowList2 > td.tdRight"));
            string price = priceWithVAT.Text;

            ws1.Cell(3, 10).Value = price;

            workbook.SaveAs(fileName);


        }

        [Test]
        public void Test4_InputIncorrectDataForSenderAndGetAlert()
        {
            var fileName = Path.Combine(Environment.CurrentDirectory, "Data\\speedy.xlsx");
            //var fileName = Path.Combine(@"C:\Users\smasl\source\repos\ReadFromExcel\ReadFromExcel\Data", "\\speedy.xlsx");
            //C:\Users\smasl\source\repos\ReadFromExcel\ReadFromExcel\Data
            var workbook = new XLWorkbook(fileName);
            var ws1 = workbook.Worksheet(1);
            int iRow = 1;
            while (!ws1.Cell(iRow, 1).IsEmpty())
            {
                var row = "";
                int iColumn = 1;
                while (!ws1.Cell(iRow, iColumn).IsEmpty())
                {
                    row = row + ws1.Cell(iRow, iColumn).Value.ToString() + ",";
                    //string tempValue = ws1.Cell(iRow, iColumn).Value.ToString();
                    //Console.OutputEncoding = Encoding.UTF8;
                    //Console.WriteLine(ws1.Cell(iRow, iColumn).Value);
                    iColumn++;
                }
                //Console.WriteLine(row);
                iRow++;
            }

            string sAddress = ws1.Cell(5, 1).Value.ToString();
            //string rAddress = ws1.Cell(3, 2).Value.ToString();

            IWebElement senderAddress = driver.FindElement(By.Id("sndrSiteName"));
            senderAddress.SendKeys(sAddress);
            System.Threading.Thread.Sleep(1500);
            senderAddress.SendKeys(Keys.Enter);

            System.Threading.Thread.Sleep(1500);
            
            IWebElement packageType = driver.FindElement(By.Id("packageType"));
            packageType.Click();

            System.Threading.Thread.Sleep(1500);

            IAlert alert = driver.SwitchTo().Alert();
            String alertcontent = alert.Text;
            Assert.AreEqual(alertcontent, "Изберете населенo място за ПОДАТЕЛ");

            alert.Accept();

        }

        [Test]
        public void Test5_InputIncorrectDataForRecieverAndGetAlert()
        {
            var fileName = Path.Combine(Environment.CurrentDirectory, "Data\\speedy.xlsx");
            var workbook = new XLWorkbook(fileName);
            var ws1 = workbook.Worksheet(1);
            int iRow = 1;
            while (!ws1.Cell(iRow, 1).IsEmpty())
            {
                var row = "";
                int iColumn = 1;
                while (!ws1.Cell(iRow, iColumn).IsEmpty())
                {
                    row = row + ws1.Cell(iRow, iColumn).Value.ToString() + ",";
                    //string tempValue = ws1.Cell(iRow, iColumn).Value.ToString();
                    //Console.OutputEncoding = Encoding.UTF8;
                    //Console.WriteLine(ws1.Cell(iRow, iColumn).Value);
                    iColumn++;
                }
                //Console.WriteLine(row);
                iRow++;
            }

            string sAddress = ws1.Cell(6, 1).Value.ToString();
            string rAddress = ws1.Cell(6, 2).Value.ToString();

            IWebElement senderAddress = driver.FindElement(By.Id("sndrSiteName"));
            senderAddress.SendKeys(sAddress);
            System.Threading.Thread.Sleep(1500);
            senderAddress.SendKeys(Keys.Enter);

            System.Threading.Thread.Sleep(1500);

            IWebElement recieverAddress = driver.FindElement(By.Id("rcptSiteName"));
            recieverAddress.SendKeys(rAddress);
            System.Threading.Thread.Sleep(1500);
            senderAddress.SendKeys(Keys.Enter);

            System.Threading.Thread.Sleep(1500);


            IWebElement packageType = driver.FindElement(By.Id("packageType"));
            packageType.Click();

            System.Threading.Thread.Sleep(1500);

            IAlert alert = driver.SwitchTo().Alert();
            String alertcontent = alert.Text;
            Assert.AreEqual(alertcontent, "Изберете населенo място за ПОЛУЧАТЕЛ");

            alert.Accept();

        }


        [TearDown]
        public void CleanUp()
        {
            driver.Close();
            Console.WriteLine("Driver closed");
        }



    }
}