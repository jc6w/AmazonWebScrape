/*********************************************************************************
*
* This program is to grab some data from the Amazon website using Selenium and
* EPPlus to write the data into separate Excel spreadsheets. This is to grab
* the autocomplete suggestions, as well as the first page results of the site.
* This also filters the search results to not include ads within the results page.
* 
* JC5044528@Syntelinc.com
*
**********************************************************************************/
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;

namespace AmazonWebScrape
{
    class Program
    {
        private static IWebDriver driver;

        //To store results
        private static List<string> autoSuggest = new List<string>();
        private static List<string> resElement = new List<string>(4);
        private static List<List<string>> searchRes = new List<List<string>>();

        //Initialize
        private static void Setup()
        {
            //Use Microsoft WebDriver for Edge
            driver = new EdgeDriver();
            
            //Maximize window
            driver.Manage().Window.Maximize();

            //Initialize waiting between elements
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
        }

        //Checks to see if the element is present in the DOM of the page by a By object
        private static bool isElementPresent(By element)
        {
            try
            {
                driver.FindElement(element);
                return true;
            }
            catch (NoSuchElementException e)
            {
                return false;
            }
        }

        //Filters search result list using certain keywords (prohibits ads)
        private static bool listFilter(string s)
        {
            if (new[] { "Sponsored", "Our Brand", "Shop by Category" }.Any(x => s.Contains(x)))
            {
                return true;
            }
            return false;
        }

        //Looks for autocomplete suggestions by Amazon
        private static void findSuggest()
        {
            //Check if the element housing autocomplete suggestions is there
            if (isElementPresent(By.Id("suggestions-template")))
            {
                //Iterate through all suggestions and add their text to the autoSuggest list
                for (int x = 0; x < 11; x++)
                {
                    IWebElement autocomp = driver.FindElement(By.Id("issDiv" + x));

                    if (autocomp.Text.Contains("in "))
                    {
                        autoSuggest.Add("To Department " + autocomp.Text);
                    }
                    else
                    {
                        autoSuggest.Add(autocomp.Text);
                    }
                }
            }
        }

        //Finds and adds all results(filtered) found to searchResults list
        private static void findResults()
        {
            //Initialize local variables
            int resultNum = 0;
            string prod = "";

            //do-while to iterate through all results in page
            do
            {
                //State/restate XPath of result element based on resultNum
                prod = "//li[@id=\"result_" + resultNum + "\"]";

                //Check if this is second iteration and thereafter
                if (resultNum > 0)
                {
                    resElement = new List<string>(4);
                }

                //Check if no more results are found
                if (!isElementPresent(By.XPath(prod)))
                {
                    //Exit do-while loop
                    break;
                }
                //If results are found
                IWebElement searchResults = driver.FindElement(By.XPath(prod));

                //Redo wait time between elements for faster processing after page has been loaded
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);
                
                //If element text contains "Sponsored", skip to next iteration
                if (listFilter(searchResults.Text))
                {
                    resultNum++;
                }
                else
                {
                    //Find product's Name/description
                    string prodName = prod + "//descendant::h2";
                    if (isElementPresent(By.XPath(prodName)))
                    {
                        searchResults = driver.FindElement(By.XPath(prodName));
                        resElement.Add(searchResults.Text);  
                    }
                    else
                    {
                        resElement.Add(null);
                    }

                    //Find product's seller
                    string prodSeller = prod + "//descendant::div[1]/div[2]/span[2]";
                    if (isElementPresent(By.XPath(prodSeller)))
                    {
                        searchResults = driver.FindElement(By.XPath(prodSeller));
                        resElement.Add(searchResults.Text);
                    }
                    else
                    {
                        resElement.Add(null);
                    }

                    //Find product's type
                    string prodType = prod + "//descendant::h3";
                    string type = "";
                    if (isElementPresent(By.XPath(prodType)))
                    {
                        searchResults = driver.FindElement(By.XPath(prodType));
                        type = searchResults.Text;
                        resElement.Add(searchResults.Text);
                    }
                    else
                    {
                        resElement.Add(null);
                    }

                    //Find product's price
                    //There are multiple lines, as I have found the location of elements to be varied at times (like those of music results)
                    string prodPrice1 = prod + "/div/div/div/div[2]/div[2]/div[1]/div[1]/a/span[1]";
                    string prodPrice2 = prod + "/div/div/div/div[2]/div[2]/div[1]/div[2]/a/span[1]";
                    string prodPrice3 = prod + "/div/div/div/div[2]/div[2]/div[1]/div[5]/div/span[1]";
                    string prodPrice4 = prod + "/div/div[2]/div/div[2]/div[2]/div[1]/div[1]/a/span[1]";
                    if (isElementPresent(By.XPath(prodPrice1)))
                    {
                        searchResults = driver.FindElement(By.XPath(prodPrice1));
                        resElement.Add(searchResults.GetAttribute("innerHTML").ToString());
                    }
                    else if (isElementPresent(By.XPath(prodPrice2)) && !(type.Contains("MP3")))
                    {
                        searchResults = driver.FindElement(By.XPath(prodPrice2));
                        resElement.Add(searchResults.GetAttribute("innerHTML").ToString());
                    }
                    else if (isElementPresent(By.XPath(prodPrice3)))
                    {
                        searchResults = driver.FindElement(By.XPath(prodPrice3));
                        resElement.Add(searchResults.GetAttribute("innerHTML").ToString());
                    }
                    else if (isElementPresent(By.XPath(prodPrice4)))
                    {
                        searchResults = driver.FindElement(By.XPath(prodPrice4));
                        resElement.Add(searchResults.GetAttribute("innerHTML").ToString());
                    }
                    else
                    {
                        resElement.Add(null);
                    }

                    searchRes.Add(resElement);
                    resultNum++;

                }

            } while (isElementPresent(By.XPath(prod)));
        }

        //Print all results to Excel
        private static bool toExcel(ExcelPackage pack)
        {
            try
            {
                //Create a worksheet within package, and name it Amazon Suggestions
                ExcelWorksheet ws = pack.Workbook.Worksheets.Add("Amazon Suggestions " + autoSuggest[0]);


                //Add all information to Excel Sheet to corresponding cells
                for (int x = 0; x < autoSuggest.Count; x++)
                {
                    ws.Cells[x + 1, 1].Value = autoSuggest[x];
                }

                //Autofit columns to fit data
                ws.Cells[ws.Dimension.Address].AutoFitColumns();

                //Create a worksheet within package, and name it Amazon Suggestions
                ws = pack.Workbook.Worksheets.Add("Amazon Search Results " + autoSuggest[0]);

                //Create headers
                ws.Cells["A1"].Value = "Product Name";
                ws.Cells["B1"].Value = "Product Seller";
                ws.Cells["C1"].Value = "Product Type";
                ws.Cells["D1"].Value = "Product Price";

                //Add all information to Excel Sheet to corresponding cells
                for (int x = 0; x < searchRes.Count; x++)
                {
                    for (int y = 0; y < searchRes[x].Count; y++)
                    {
                        ws.Cells[x + 2, y + 1].Value = searchRes[x][y];
                    }
                }

                //Autofit columns to fit data
                ws.Cells[ws.Dimension.Address].AutoFitColumns();

                return true;
            }
            catch (Exception e)
            {
                return false;
            }
        }

        //Main Function
        public static void Main(string[] args)
        {
            //From EPPlus
            //Create a new package to create an excel file
            ExcelPackage pack = new ExcelPackage();

            //Where to store the excel file
            FileInfo fileName = new FileInfo("C:/Users/JC5044528/Desktop/Amazon.xlsx");

            //Initialize and start driver
            Setup();

            //Either command works to navigate to specified URL
            driver.Url = "www.amazon.com";
            //driver.Navigate().GoToUrl("www.amazon.com");

            //Find Amazon's search bar
            IWebElement searchBox = driver.FindElement(By.Id("twotabsearchtextbox"));

            //Search for a specific object. The ImplicitWait is there to make the browser wait for the pop-up of suggestions to appear
            searchBox.SendKeys("USB C Cable");

            //Finds the autocomplete suggestions on the pulldown
            findSuggest();

            //This is to go and find actual results on Amazon
            driver.FindElement(By.ClassName("nav-input")).Click();

            //Goes through the first result page
            findResults();

            //New worksheets to print to Excel
            toExcel(pack);

            //Save to a new file
            pack.SaveAs(fileName);

            //Close browser
            driver.Close();

            //Close driver
            driver.Quit();

            //These are all me just testing different parts of the page and driver commands
            /* 
            driver.FindElement(By.Id("pagnNextLink")).Click();
            
            for (int x = 0; x < 20; x++)
            {
                string searchString = "//*[@id, \"result_" + x + "\"]/div/div/div/div[2]";
                IWebElement searchRes = driver.FindElement(By.XPath(searchString));
                if (searchString.Contains("Sponsored"))
                {
                    continue;
                }
                res.Add(searchRes.Text);
            }
            driver.Navigate().GoToUrl("https:www.google.com");
            IWebElement element = driver.FindElement(By.XPath("//input[@name='q']"));
            element.SendKeys("Syntel Inc");*/
        }
    }
}
