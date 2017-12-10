﻿using System;
using System.Text.RegularExpressions;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Jitbit.Utils;
using System.IO;
using System.Text;

namespace AliOrderScraper
{
    class Program
    {
        static void Main(string[] args)
        {
            ScrapeAliExpressOrders("", "AliExpressOrders");
        }

        static void ScrapeAliExpressOrders(string path, string fileNamePrefix)
        {
            // http://blog.hanxiaogang.com/2017-07-29-aliexpress/
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("user-data-dir=C:/Users/pnerseth/AppData/Local/Google/Chrome/User Data/Default");
            options.AddArguments("--start-maximized");
            IWebDriver driver = new ChromeDriver(options);
            driver.Navigate().GoToUrl("https://login.aliexpress.com");

            try
            {
                var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                wait.Until(ExpectedConditions.UrlToBe("https://www.aliexpress.com/"));
            }
            catch (WebDriverTimeoutException)
            {
                Console.WriteLine("Timeout - Logged in to AliExpress to late. Stopping.");
                return;
            }

            // go to order list
            driver.Navigate().GoToUrl("https://trade.aliexpress.com/orderList.htm");

            // identify how many pages on order page (1/20)
            var tuple = GetAliExpressOrderPageNumber(driver);
            int curPage = tuple.Item1;
            int numPages = tuple.Item2;
            Console.WriteLine("Found {0} Pages", numPages);

            // scrape one and one page and store in csv writer object
            var myExport = new CsvExport();
            for (int i = 1; i <= numPages; i++) {
                if (ScrapeAliExpressOrderPage(myExport, driver, i))
                {
                    Console.WriteLine("Successfully Scraped Order Page {0}", i);
                }
            }

            // set export filename
            string exportFilePath = Path.Combine(path, string.Format("{0}-{1:yyyy-MM-dd_HH-mm}.csv", fileNamePrefix, DateTime.Now));
            myExport.ExportToFile(exportFilePath, true);
        }

        static bool ScrapeAliExpressOrderPage(CsvExport myExport, IWebDriver driver, int curPage)
        {
            // change page
            driver.FindElement(By.CssSelector("input[id$='gotoPageNum']")).SendKeys(curPage.ToString());
            driver.FindElement(By.XPath("//*[@id='btnGotoPageNum']")).Click();

            // check that we in fact got to the right page
            var tuple = GetAliExpressOrderPageNumber(driver);
            int newPage = tuple.Item1;
            int numPages = tuple.Item2;

            if (curPage != newPage) return false;

            // scrape
            Console.WriteLine("Reading Page {0} of {1} Pages", curPage, numPages);
            ScrapeAliExpressOrderPageEntry(myExport, driver, curPage);

            return true;
        }

        static void ScrapeAliExpressOrderPageEntry(CsvExport myExport, IWebDriver driver, int curPage)
        {
            var orderEntries = driver.FindElements(By.XPath("//tbody[contains(@class, 'order-item-wraper ')]"));
            Console.WriteLine("Found {0} orders on page {1}", orderEntries.Count, curPage);

            int pageCount = 1;
            foreach (var orderEntry in orderEntries)
            {
                Console.WriteLine("Reading order number {0}", pageCount++);

                var orderId = orderEntry.FindElement(By.XPath("tr[@class='order-head']/td[@class='order-info']/p[@class='first-row']/span[@class='info-body']")).Text;
                var orderTime = orderEntry.FindElement(By.XPath("tr[@class='order-head']/td[@class='order-info']/p[@class='second-row']/span[@class='info-body']")).Text;

                var storeName = orderEntry.FindElement(By.XPath("tr[@class='order-head']/td[@class='store-info']/p[@class='first-row']/span[@class='info-body']")).Text;
                var storeUrl = orderEntry.FindElement(By.XPath("tr[@class='order-head']/td[@class='store-info']/p[@class='second-row']/a")).GetAttribute("href");

                var orderAmount = orderEntry.FindElement(By.XPath("tr[@class='order-head']/td[@class='order-amount']/div[@class='amount-body']/p[@class='amount-num']")).Text;

                Console.WriteLine("Order no {0} was ordered on the {1}\n{2}, {3}, {4}", orderId, orderTime, storeName, storeUrl, orderAmount);

                myExport.AddRow();
                myExport["OrderId"] = orderId;
                myExport["OrderTime"] = orderTime;
                myExport["StoreName"] = storeName;
                myExport["StoreUrl"] = storeUrl;
                myExport["OrderAmount"] = orderAmount;

                // for each order line
                var orderLines = orderEntry.FindElements(By.XPath("tr[@class='order-body']"));

                int orderLineCount = 1;
                bool first = true;
                var builder = new StringBuilder(); // initially empty
                foreach (var orderLine in orderLines) {
                    // append newline after each line
                    if (first) {
                        first = false;
                    } else {
                        builder.Append("\n"); 
                    }

                    var productTitleElement = orderLine.FindElement(By.XPath("td[@class='product-sets']/div[@class='product-right']/p[@class='product-title']/a")); ;
                    var productId = productTitleElement.GetAttribute("productId");
                    var productTitle = productTitleElement.Text;

                    var productAmount = orderLine.FindElement(By.XPath("td[@class='product-sets']/div[@class='product-right']/p[@class='product-amount']")).Text;
                    var productProperty = orderLine.FindElement(By.XPath("td[@class='product-sets']/div[@class='product-right']/p[@class='product-property']")).Text;

                    Console.WriteLine("{0}. [{1}] {2}\n{3} {4}", orderLineCount, productId, productTitle, productProperty, productAmount);
                    builder.AppendFormat("{0}. [{1}] {2}\n{3} {4}", orderLineCount, productId, productTitle, productProperty, productAmount);
                    orderLineCount++;
                }
                myExport["OrderLines"] = builder.ToString();

                // read order contact information (buyer)
                GetAliExpressContactFromOrder(myExport, driver, orderId);

                // new line
                Console.WriteLine();
            }
        }

        static void GetAliExpressContactFromOrder(CsvExport myExport, IWebDriver driver, string orderId)
        {
            // https://trade.aliexpress.com/order_detail.htm?orderId=81495464493633

            // open a new tab and set the context
            var chromeDriver = (ChromeDriver)driver;

            // save a reference to our original tab's window handle
            var originalTabInstance = chromeDriver.CurrentWindowHandle;

            // execute some JavaScript to open a new window
            chromeDriver.ExecuteScript("window.open();");

            // save a reference to our new tab's window handle, this would be the last entry in the WindowHandles collection
            var newTabInstance = chromeDriver.WindowHandles[driver.WindowHandles.Count - 1];

            // switch our WebDriver to the new tab's window handle
            chromeDriver.SwitchTo().Window(newTabInstance);

            // lets navigate to a web site in our new tab
            string url = String.Format("https://trade.aliexpress.com/order_detail.htm?orderId={0}", orderId);
            driver.Navigate().GoToUrl(url);

            // find contact information
            // example:
            // <li><label> Contact Name :</label><span i18entitle = 'Contact Name' class="i18ncopy">Reidar Krogsaeter</span>
            // find the span element with correct i18entitle contained within the li that contains a label whose string value contains the substring Contact Name
            string contactName = driver.FindElement(By.XPath("//li[label[contains(., 'Contact Name')]]/span[@i18entitle='Contact Name']")).Text;
            string contactAddress = driver.FindElement(By.XPath("//li[label[contains(., 'Address')]]/span[@i18entitle='Address']")).Text;
            string contactZipCode = driver.FindElement(By.XPath("//li[label[contains(., 'Zip Code')]]/span[@i18entitle='Zip Code']")).Text;
            // example 2:
            // find the first following li element after the address and extract the span that has class i18ncopy
            string contactAddress2 = driver.FindElement(By.XPath("//li[label[contains(., 'Address')]]/following-sibling::li[1]/span[@class='i18ncopy']")).Text;

            Console.WriteLine("Contact {0} {1} {2} {3}", contactName, contactAddress, contactAddress2, contactZipCode);

            myExport["ContactName"] = contactName;
            myExport["ContactAddress"] = contactAddress;
            myExport["ContactAddress2"] = contactAddress2;
            myExport["ContactZipCode"] = contactZipCode;

            // now lets close our new tab
            chromeDriver.ExecuteScript("window.close();");

            // and switch our WebDriver back to the original tab's window handle
            chromeDriver.SwitchTo().Window(originalTabInstance);

            // and have our WebDriver focus on the main document in the page to send commands to 
            chromeDriver.SwitchTo().DefaultContent();
        }

        static Tuple<int, int> GetAliExpressOrderPageNumber(IWebDriver driver)
        {
            int curPage = 0;
            int numPages = 0;
            var pageLabel = driver.FindElement(By.CssSelector("label[class$='ui-label']")).Text;
            Regex regexObj = new Regex(@"(\d+)/(\d+)", RegexOptions.IgnoreCase);
            Match matchResults = regexObj.Match(pageLabel);
            if (matchResults.Success)
            {
                curPage = int.Parse(matchResults.Groups[1].Value);
                numPages = int.Parse(matchResults.Groups[2].Value);
                return new Tuple<int, int>(curPage, numPages);
            }

            return null;
        }
    }
}
