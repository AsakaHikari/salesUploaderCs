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
using ClosedXML.Excel;
using MailKit.Security;

namespace mgs
{
    class Program
    {

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Google Sheets API .NET Quickstart";
        static void Main(string[] args)
        {


            Program pg = new Program();
            //pg.Mgs();
            pg.Unext();
        }

        public void Unext()
        {
            string download = System.Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "/Downloads";

            // 対象ファイルを検索する
            string[] fileList = Directory.GetFileSystemEntries(download, @"hol0001928*.xlsx");
            
            // 抽出したファイル数を出力
            //Console.WriteLine("file num = " + fileList.Length.ToString());
            foreach (string path in fileList)
            {
                File.Delete(path);
            }
            IWebDriver driver = new ChromeDriver();

            var wait = new WebDriverWait(driver, new TimeSpan(0, 0, 20));

            driver.Url = "https://rbivy.eco-serv.jp/unext/index/";
            wait.Until(
                ExpectedConditions.ElementExists(By.Name("loginId"))
            );
            IWebElement text = driver.FindElement(By.Name("loginId"));
            text.SendKeys("hol0001928");
            text = driver.FindElement(By.Name("passwd"));
            text.SendKeys("hoge0794");
            System.Threading.Thread.Sleep(2000);
            IWebElement btn = driver.FindElement(By.CssSelector("input[type='submit']"));

            btn.Click();

            
            wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("a[href='/unext/bapPublishedBillAppSearch/']")));
            btn = driver.FindElement(By.CssSelector("a[href='/unext/bapPublishedBillAppSearch/']"));
            btn.Click();
             wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("a[class='wb_pc_cl_user_bill_dl_lnk pdf link_text icon_attach']")));
            btn = driver.FindElement(By.CssSelector("a[class='wb_pc_cl_user_bill_dl_lnk pdf link_text icon_attach']"));
            btn.Click();
            
            // 対象ファイルを検索する
            fileList = Directory.GetFileSystemEntries(download, @"hol0001928*.xlsx");

            string filePath;
            long filesize1, filesize2;
            do
            {
                Thread.Sleep(500);
                fileList = Directory.GetFileSystemEntries(download, @"hol0001928*.xlsx");
                filePath = fileList[0];
            } while (fileList.Length == 0 || !File.Exists(filePath));
            do
            {

                filesize1 = (new System.IO.FileInfo(filePath)).Length;  // check file size
                Thread.Sleep(500);      // wait for 5 seconds
                filesize2 = (new System.IO.FileInfo(filePath)).Length;  // check file size again

            } while (filesize1 != filesize2);

            XLWorkbook workbook = new XLWorkbook(filePath);
            IXLWorksheet sheet = workbook.Worksheets.Worksheet(3);
            string str = "";
            for (int i=4;i<sheet.LastRowUsed().RowNumber()+1;i++)
            {
                str+=(sheet.Cell(i,2).Value+","+ sheet.Cell(i, 8).Value+"\n");
            }
            string fileName = Path.GetFileNameWithoutExtension(filePath);
            FileStream fs = new FileStream("./"+fileName+".csv", FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);
            sw.WriteLine(str);
            sw.Close();
            fs.Close();
            Mail("pa01m97ksuxa615@gmail.com", "xscBLjNyHmDb9yX", "pa01m97ksuxa615@gmail.com", "nex@m-trax.net", fileName+".csv", str, "./" + fileName + ".csv");
        }

        public void SpreadSheet()
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
            String spreadsheetId = "13XJSI6SRQ_yClthHD9tLTw2XlzcjQEz7n1ubsr9ZrBM";
            String range = "Class Data!A2:E";

            IList<IList<Object>> values = new List<IList<object>>();
            IList<Object> list1 = new List<object>();
            list1.Add("abcde");
            values.Add(list1);
            ValueRange body = new ValueRange();
            body.Values = values;

            SpreadsheetsResource.ValuesResource.UpdateRequest ur = service.Spreadsheets.Values.Update(body, spreadsheetId, range);
            ur.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
            UpdateValuesResponse result = ur.Execute();

            Console.WriteLine("%d cells updated.", result.UpdatedCells);

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
            for (int i = 1; i < trs.Count; i++)
            {
                var tr = trs[i];
                var tds = tr.FindElements(By.CssSelector("td:not([rowspan])"));
                if (PCSP == 0)
                {
                    if (tds[0].Text.Equals("PC Total"))
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
            String param = "";
            foreach (KeyValuePair<string, int> pair in dicPC)
            {
                param = param + (pair.Key + ":" + pair.Value + ",");
            }
            Console.WriteLine();
            foreach (KeyValuePair<string, int> pair in dicSP)
            {
                param = param + (pair.Key + ":" + pair.Value + ",");
            }
            //driver.Url = "https://docs.google.com/spreadsheets/d/13XJSI6SRQ_yClthHD9tLTw2XlzcjQEz7n1ubsr9ZrBM/edit#gid=11310634";
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

        public static void Mail(string id,string pass,string from,string to,string subject,string body,string path)
        {
            //Console.WriteLine("Hello SMTP World!");
            /*
            string id = "<gmailのログインID>";
            string pass = "<gmailのパスワード>";
            string from = "<宛先>";
            string to = "<自分のメール>";
            string subject = "送信テスト : " + DateTime.Now.ToString();
            string body = "from t.masuda";
            */
#if false
    var smtp = new System.Net.Mail.SmtpClient();
    smtp.Host = "smtp.gmail.com"; //SMTPサーバ
    smtp.Port = 587;              //SMTPポート
    smtp.EnableSsl = true;
    smtp.Credentials = new System.Net.NetworkCredential(id, pass); //認証
    var msg  = new System.Net.Mail.MailMessage(from, to, subject, body);
    smtp.Send(msg); //メール送信
#else

            var smtp = new MailKit.Net.Smtp.SmtpClient();
            smtp.Connect("smtp.gmail.com", 587, SecureSocketOptions.Auto);
            smtp.Authenticate(id, pass);

            var mail = new MimeKit.MimeMessage();
            var builder = new MimeKit.BodyBuilder();

            mail.From.Add(new MimeKit.MailboxAddress("", from));
            mail.To.Add(new MimeKit.MailboxAddress("", to));
            mail.Subject = subject;
            builder.TextBody = body;
            //mail.Body = builder.ToMessageBody();
            
            //var path = @"C:\Windows\Web\Wallpaper\Theme2\img10.jpg"; // 添付したいファイル
            var attachment = new MimeKit.MimePart("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                Content = new MimeKit.MimeContent(File.OpenRead(path)),
                ContentDisposition = new MimeKit.ContentDisposition(),
                ContentTransferEncoding = MimeKit.ContentEncoding.Base64,
                FileName = System.IO.Path.GetFileName(path)
            };

            var multipart = new MimeKit.Multipart("mixed");
            multipart.Add(builder.ToMessageBody());
            multipart.Add(attachment);

            mail.Body = multipart;

            smtp.Send(mail);
            smtp.Disconnect(true);
#endif

            Console.WriteLine("メールを送信しました");
        }
    }



}

