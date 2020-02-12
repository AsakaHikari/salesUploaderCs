using System;
using System.Collections.Generic;
using System.Net;
using HtmlAgilityPack;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.IO;
using System.Threading;

namespace mgs
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
        static string ApplicationName = "Google Sheets API .NET Quickstart";
        static void Main(string[] args)
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Sheets API service.
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Define request parameters.
            String spreadsheetId = "1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms";
            String range = "Class Data!A2:E";
            IList<IList<Object>> values = new List<IList<object>>();
            IList<Object> list1 = new List<object>();
            list1.Add("abcde");
            values.Add(list1);
            ValueRange body = new ValueRange();
            body.Values = values;
            UpdateValuesResponse result =
                    service.Spreadsheets.Values.Update(body, spreadsheetId, range)
                            .ValueInputOption = valueInputOption;
                            .execute();
            System.out.printf("%d cells updated.", result.getUpdatedCells());
            Program pg = new Program();
            //pg.Mgs();

        }

        public void Mgs()
        {
            IWebDriver driver = new ChromeDriver();

            var wait = new WebDriverWait(driver, new TimeSpan(0, 0, 20));

            driver.Url = "http://kumanekohonnpo:2ZP911MO@www.mgstage.com/makeradmin/?r=sales";


            wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("a[href='?r=Product/index']")));
            IWebElement btn = driver.FindElement(By.CssSelector("a[href='?r=Product/index']"));
            btn.Click();
            IWebElement text = driver.FindElement(By.Name("type"));
            text.SendKeys("1");
            IWebElement radioButton = driver.FindElement(By.CssSelector("input[value='1']"));
            radioButton.Click();
            text = driver.FindElement(By.Name("sh[begin_date]"));
            text.SendKeys("2019/12/01");
            text = driver.FindElement(By.Name("sh[end_date]"));
            text.SendKeys("2019/12/31");
            System.Threading.Thread.Sleep(2000);
            btn = driver.FindElement(By.CssSelector("a[onclick='search_submit()']"));

            btn.Click();
                wait.Until(
                            ExpectedConditions.ElementExists(By.Name("viewport"))
            );
            IWebElement table = driver.FindElement(By.CssSelector("table[class='m-table2']"));
            var trs = table.FindElements(By.TagName("tr"));
            int PCSP = 0;
            var dicPC = new Dictionary<String, int>();
            var dicSP = new Dictionary<String, int>();
            for (int i=1;i<trs.Count;i++){
                var tr = trs[i];
                var tds=tr.FindElements(By.CssSelector("td:not([rowspan])"));
                if (PCSP == 0)
                {
                    if(tds[0].Text.Equals("PC Total"))
                    {
                        PCSP = 1;
                        continue;
                    }
                    int value = int.Parse(tds[3].Text);
                    if (dicPC.ContainsKey(tds[0].Text))
                    {
                        value += dicPC[tds[0].Text];
                        dicPC.Remove(tds[0].Text);
                        dicPC.Add(tds[0].Text, value);
                    }
                    else
                    {
                        dicPC.Add(tds[0].Text, value);
                    }
                }
                else
                {

                    if (tds[0].Text.Equals("SP Total"))
                    {
                        PCSP = 2;
                        break;
                    }
                    int value = int.Parse(tds[3].Text);
                    if (dicSP.ContainsKey(tds[0].Text))
                    {
                        value += dicSP[tds[0].Text];
                        dicSP.Remove(tds[0].Text);
                        dicSP.Add(tds[0].Text, value);
                    }
                    else
                    {
                        dicSP.Add(tds[0].Text, value);
                    }
                    
                }
            }
            foreach(KeyValuePair<string, int>pair in dicPC)
            {
                Console.WriteLine(pair.Key + ":" + pair.Value);
            }
            Console.WriteLine();
            foreach (KeyValuePair<string, int> pair in dicSP)
            {
                Console.WriteLine(pair.Key + ":" + pair.Value);
            }
            driver.Url = "https://script.google.com/macros/s/AKfycbzVV0n-CRBwNAUuEVmUxpfMpf4EbScClGAKWdFSgDQccLmsMj0j/exec?param=fuga";
        }

        public void Fanza()
        {
            IWebDriver driver = new ChromeDriver();

            var wait = new WebDriverWait(driver, new TimeSpan(0, 0, 20));

            driver.Url = "https://owner-admin.dmm.com/owner";
            wait.Until(
                ExpectedConditions.ElementExists(By.Name("mail"))
            );
            IWebElement text = driver.FindElement(By.Name("mail"));
            text.SendKeys("nex@m-trax.net");
            text = driver.FindElement(By.Name("password"));
            text.SendKeys("hoge0794");
            System.Threading.Thread.Sleep(2000);
            IWebElement btn = driver.FindElement(By.CssSelector("button[type='submit']"));

            btn.Click();


            wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("a[href='https://owner-admin.dmm.com/owner/search/product']")));
            btn = driver.FindElement(By.CssSelector("a[href='https://owner-admin.dmm.com/owner/search/product']"));
            btn.Click();

            wait.Until(
                ExpectedConditions.ElementExists(By.Name("dt_year_from"))
            );
            new SelectElement(driver.FindElement(By.Name("dt_year_from"))).SelectByValue("2019");
            new SelectElement(driver.FindElement(By.Name("dt_month_from"))).SelectByValue("12");
            new SelectElement(driver.FindElement(By.Name("dt_day_from"))).SelectByValue("1");
            new SelectElement(driver.FindElement(By.Name("dt_year_to"))).SelectByValue("2019");
            new SelectElement(driver.FindElement(By.Name("dt_month_to"))).SelectByValue("12");
            new SelectElement(driver.FindElement(By.Name("dt_day_to"))).SelectByValue("31");

            btn = driver.FindElement(By.CssSelector("button[onClick='SearchProductSubmit(this)']"));
            wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("button[onClick='SearchProductSubmit(this)']")));
            ExecuteJavaScript(driver, "window.scrollTo(0, 500);");
            System.Threading.Thread.Sleep(1000);
            btn.Click();
            wait.Until(
                ExpectedConditions.ElementExists(By.Name("_token"))
            );
            System.Threading.Thread.Sleep(10000);
            ExecuteJavaScript(driver, "CSVExpload(this, '1')");
        }

        public static void ExecuteJavaScript(IWebDriver driver, string script)
        {
            if (driver is IJavaScriptExecutor)
                ((IJavaScriptExecutor)driver).ExecuteScript(script);
            else
                throw new WebDriverException();
        }
    }

}
