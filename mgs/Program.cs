using System;
using System.Collections.Generic;
using System.Collections;
using System.Net;

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
using DocumentFormat.OpenXml.Office2013.PowerPoint.Roaming;

using System.Text;
using DocumentFormat.OpenXml.Vml;
using System.Configuration;
using DocumentFormat.OpenXml.Wordprocessing;

namespace mgs
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = { SheetsService.Scope.Spreadsheets ,Google.Apis.Drive.v3.DriveService.Scope.Drive};
        static string ApplicationName = "Google Sheets API .NET Quickstart";

        private void sheets()
        {
            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.ReadWrite))
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

            String spreadsheetId = ConfigurationManager.AppSettings.Get("SpreadSheetID");
            IList<IList<Object>> values;
            IList<IList<Object>> valueswrite;
            String range;
            ValueRange response;
            ValueRange body;
            String rangewrite;
            SpreadsheetsResource.ValuesResource.GetRequest request;
            SpreadsheetsResource.ValuesResource.UpdateRequest ur;
            UpdateValuesResponse result;

            if (ConfigurationManager.AppSettings.Get("U-NEXT").Equals("true"))
            {

                range = ConfigurationManager.AppSettings.Get("U-NEXT_sheetname") + "!A1:"+ ConfigurationManager.AppSettings.Get("MaxColumn");
                request = service.Spreadsheets.Values.Get(spreadsheetId, range);

                response = request.Execute();

                values = response.Values;
                valueswrite = new List<IList<Object>>();
                //ここで売上報告を受け取る
                
                this.Unext(values, valueswrite);

                body = new ValueRange();
                body.Values = valueswrite;
                rangewrite = ConfigurationManager.AppSettings.Get("U-NEXT_sheetname") + "!A1:" + ConfigurationManager.AppSettings.Get("MaxColumn");


                ur = service.Spreadsheets.Values.Update(body, spreadsheetId, rangewrite);
                ur.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                result=ur.Execute();
            }

            //----FANZA----
            if (ConfigurationManager.AppSettings.Get("FANZA").Equals("true"))
            {
                range = ConfigurationManager.AppSettings.Get("FANZA_sheetname") + "!A1:" + ConfigurationManager.AppSettings.Get("MaxColumn");
                request =
                        service.Spreadsheets.Values.Get(spreadsheetId, range);

                response = request.Execute();
                values = response.Values;
                valueswrite = new List<IList<Object>>();

                this.Fanza(values, valueswrite);

                body = new ValueRange();
                body.Values = valueswrite;
                rangewrite = ConfigurationManager.AppSettings.Get("FANZA_sheetname") + "!A1:" + ConfigurationManager.AppSettings.Get("MaxColumn");

                ur = service.Spreadsheets.Values.Update(body, spreadsheetId, rangewrite);
                ur.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                result = ur.Execute();
            }

            //----MGS----
            if (ConfigurationManager.AppSettings.Get("MGS").Equals("true"))
            {
                range = ConfigurationManager.AppSettings.Get("MGS_sheetname") + "!A1:" + ConfigurationManager.AppSettings.Get("MaxColumn");
                request =
                        service.Spreadsheets.Values.Get(spreadsheetId, range);

                response = request.Execute();
                values = response.Values;
                valueswrite = new List<IList<Object>>();

                this.Mgs(values, valueswrite);

                body = new ValueRange();
                body.Values = valueswrite;
                rangewrite = ConfigurationManager.AppSettings.Get("MGS_sheetname") + "!A1:" + ConfigurationManager.AppSettings.Get("MaxColumn");

                ur = service.Spreadsheets.Values.Update(body, spreadsheetId, rangewrite);
                ur.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                result = ur.Execute();
            }

            //----MPO----
            if (ConfigurationManager.AppSettings.Get("mpo").Equals("true"))
            {
                range = ConfigurationManager.AppSettings.Get("mpo_sheetname") + "!A1:" + ConfigurationManager.AppSettings.Get("MaxColumn");
                request =
                        service.Spreadsheets.Values.Get(spreadsheetId, range);

                response = request.Execute();
                values = response.Values;
                valueswrite = new List<IList<Object>>();

                this.Mpo(values, valueswrite);

                body = new ValueRange();
                body.Values = valueswrite;
                rangewrite = ConfigurationManager.AppSettings.Get("mpo_sheetname") + "!A1:" + ConfigurationManager.AppSettings.Get("MaxColumn");

                ur = service.Spreadsheets.Values.Update(body, spreadsheetId, rangewrite);
                ur.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.USERENTERED;
                result = ur.Execute();
            }
            Console.Read();
        }
    

    static void Main(string[] args)
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            Program pg = new Program();
            pg.sheets();
        }

        public void Unext(IList<IList<Object>> values, IList<IList<Object>> valuesWrite)
        {
            DateTime targetdate = DateTime.Now.AddMonths(-int.Parse(ConfigurationManager.AppSettings.Get("U-NEXT_monthlate")));
            int year = targetdate.Year;
            int month = targetdate.Month;
            string monthstr = month.ToString("D2");

            string userid = ConfigurationManager.AppSettings.Get("U-NEXT_userid");
            string password = ConfigurationManager.AppSettings.Get("U-NEXT_password");

            string download = System.Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "/Downloads";

            // 対象ファイルを検索する
            string[] fileList = Directory.GetFileSystemEntries(download, @userid+@"*.xlsx");

            // 抽出したファイル数を出力
            //Console.WriteLine("file num = " + fileList.Length.ToString());
            foreach (string path in fileList)
            {
                File.Delete(path);
            }
            IWebDriver driver=null;
            try
            {
                driver = new ChromeDriver();
            }catch(Exception e)
            {
                Console.WriteLine("\n********* AN ERROR HAS OCCURED *********");
                Console.WriteLine("Maybe chromedriver or GoogleChrome must be updated?");
                Console.WriteLine(e.Message);
                Console.ReadKey();
                
            }

            var wait = new WebDriverWait(driver, new TimeSpan(0, 0, 20));

            driver.Url = "https://rbivy.eco-serv.jp/unext/index/";
            wait.Until(
                ExpectedConditions.ElementExists(By.Name("loginId"))
            );
            IWebElement text = driver.FindElement(By.Name("loginId"));
            text.SendKeys(userid);
            text = driver.FindElement(By.Name("passwd"));
            text.SendKeys(password);
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
            // 
            fileList = Directory.GetFileSystemEntries(download, @userid+@"*.xlsx");

            string filePath="";
            long filesize1, filesize2;
            do
            {
                Thread.Sleep(500);
                fileList = Directory.GetFileSystemEntries(download, @userid+@"*.xlsx");
                if(fileList.Length >0)filePath = fileList[0];
            } while (fileList.Length == 0 || !File.Exists(filePath));
            do
            {

                filesize1 = (new System.IO.FileInfo(filePath)).Length;  // check file size
                Thread.Sleep(500);      // wait for 5 seconds
                filesize2 = (new System.IO.FileInfo(filePath)).Length;  // check file size again

            } while (filesize1 != filesize2);

            /**
             ここに、データをスプレッドシートに書き込む処理を書く。
            filePath ... ダウンロードしたxlsxファイルのパスを表す。これに必要なデータはすべて入っています。
            xlsxを読み込んで扱える状態にするには、別にAPIが必要だと思います。
             */

            // Excelファイルを開く
            var workbook = new XLWorkbook(filePath);
            // ワークシートを取得する
            var worksheet = workbook.Worksheet("作品一覧");

            string date_ = (string)values[2][1];
            int column = 2;
            if (date_.Length >= 6)
            {
                int year_ = int.Parse(date_.Substring(0, 4));
                int month_ = int.Parse(date_.Substring(4, 2));
                column = (year - year_) * 12 + (month - month_) + 2;
                Console.WriteLine("date:" + date_ + " year:" + year_ + " month:" + month_ + " column:" + column);

            }
            addToList<IList<Object>>(valuesWrite, new List<Object>(), new List<Object>(), column);
            addToList<Object>(valuesWrite[column], year + "" + monthstr + "売上", null, 1);

            var dic = new Dictionary<String, int>();

            for (int i=4;i<=worksheet.RowCount();i++)
            {

                //foreach (string e in elms) Console.WriteLine(e);
                string id = worksheet.Cell(i, 2).GetValue<string>();
                string price = worksheet.Cell(i, 8).GetValue<string>();
                Console.WriteLine(id + "," + price);
                int value = 0;
                try
                {
                    value = int.Parse(price.Replace("\\", ""));
                }catch(Exception e)
                {
                    break;
                }
                if (dic.ContainsKey(id))
                {
                    value += dic[id];
                    dic.Remove(id);
                    dic.Add(id, value);
                }
                else
                {
                    dic.Add(id, value);
                }
            }
            int index = 0;
            foreach (KeyValuePair<string, int> pair in dic)
            {
                index++;
                //id,attach
                string id = pair.Key;
                bool flag = false;
                for (int i = 0; i < values[0].Count; i++)
                {
                    //Console.WriteLine((string)values[0][i]);
                    if (((string)values[0][i]).Contains(id.ToLower().Replace("-","")))
                    {
                        Console.WriteLine(pair.Key + " = " + values[0][i] + "(" + i + ")");
                        addToList<Object>(valuesWrite[column], pair.Value, "0", i);
                        flag = true;
                        break;
                    }
                }
                if (!flag)
                {
                    Console.WriteLine(pair.Key + " Not Found");
                }
            }
        }

        public IList<T> addToList<T>(IList<T> list,T e,T filler,int index)
        {
            if (list.Count <= index)
            {
                for(int i = list.Count; i <= index; i++)
                {
                    list.Add(filler);
                }
            }
            //Console.WriteLine(list.Count);
            list[index] = e;
            return list;
        }


        public void Mgs(IList<IList<Object>> values,IList<IList<Object>> valuesWrite)

        {
            //IList<IList<Object>> valuesWrite = new List<IList<Object>>();
            DateTime targetdate = DateTime.Now.AddMonths(-int.Parse(ConfigurationManager.AppSettings.Get("MGS_monthlate")));
            int year = targetdate.Year;
            int month = targetdate.Month;
            string monthstr = month.ToString("D2");

            string userid = ConfigurationManager.AppSettings.Get("MGS_userid");
            string password = ConfigurationManager.AppSettings.Get("MGS_password");

            string download = System.Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "/Downloads";

            IWebDriver driver = new ChromeDriver();

            var wait = new WebDriverWait(driver, new TimeSpan(0, 0, 20));

            driver.Url = "http://"+userid+":"+password+"@www.mgstage.com/makeradmin/?r=sales";


            wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("a[href='?r=Product/index']")));
            IWebElement btn = driver.FindElement(By.CssSelector("a[href='?r=Product/index']"));
            btn.Click();
            IWebElement text = driver.FindElement(By.Name("type"));
            text.SendKeys("1");
            IWebElement radioButton = driver.FindElement(By.CssSelector("input[value='1']"));
            radioButton.Click();
            text = driver.FindElement(By.Name("sh[begin_date]"));
            text.SendKeys(year+"/"+monthstr+"/01");
            text = driver.FindElement(By.Name("sh[end_date]"));
            text.SendKeys(year+"/" + monthstr + "/"+DateTime.DaysInMonth(year,month));
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
            string date_ = (string)values[2][1];
            int column = 2;
            if (date_.Length >= 6)
            {
                int year_ = int.Parse(date_.Substring(0, 4));
                int month_ = int.Parse(date_.Substring(4, 2));
                column = (year - year_) * 12 + (month - month_) + 2;
                Console.WriteLine("date:" + date_ + " year:" + year_ + " month:" + month_ + " column:" + column);

            }
            addToList<IList<Object>>(valuesWrite, new List<Object>(), new List<Object>(), column);
            addToList<Object>(valuesWrite[column], year + "" + monthstr + "売上", null, 1);

            int index = 0;
            foreach (KeyValuePair<string, int> pair in dicPC)
            {
                index++;
                //id,attach
                string id = pair.Key;
                for(int i=0;i<id.Length;i++)
                {
                    if (id[i] < 48 || id[i] > 57)
                    {
                        id = id.Substring(i).ToLower().Replace("-","");
                        break;
                    }
                }
                //Console.WriteLine(id);
                bool flag = false; 
                for(int i = 0; i < values[0].Count; i++)
                {
                    //Console.WriteLine((string)values[0][i]);
                    if (((string)values[0][i]).Contains(id))
                    {
                        Console.WriteLine(pair.Key + " = " + values[0][i] + "(" + i + ")");
                        addToList<Object>(valuesWrite[column], pair.Value, "0", i);
                        flag = true;
                        break;
                    }
                }
                if (!flag)
                {
                    Console.WriteLine(pair.Key + " Not Found");
                }
                //valuesWrite[0].Add(pair.Key);
                //valuesWrite[1].Add(pair.Value);

            }
            foreach (KeyValuePair<string, int> pair in dicSP)
            {
                index++;
                //id,attach
                string id = pair.Key.Substring(3);
                for (int i = 0; i < id.Length; i++)
                {
                    if (id[i] < 48 || id[i] > 57)
                    {
                        id = id.Substring(i).ToLower().Replace("-", "");
                        break;
                    }
                }
                //Console.WriteLine(id);
                bool flag = false;
                for (int i = 0; i < values[0].Count; i++)
                {
                    //Console.WriteLine((string)values[0][i]);
                    if (((string)values[0][i]).Contains(id))
                    {
                        Console.WriteLine(pair.Key + " = " + values[0][i] + "(" + i + ")");
                        addToList<Object>(valuesWrite[column], pair.Value, "0", i+1);
                        flag = true;
                        break;
                    }
                }
                if (!flag)
                {
                    Console.WriteLine(pair.Key + " Not Found");
                }
                
                }
            //valuesWrite[0].Add(pair.Key);
            //valuesWrite[1].Add(pair.Value);
            foreach (IList<Object> l in valuesWrite)
            {
                if (l != null)
                {
                    foreach (Object o in l)
                    {
                        Console.Write(o + " ,");
                    }
                    
                }
                Console.WriteLine("\\n");

            }

            /**
             * DicPCにはPC Totalというデータが、DicSPにはSP Totalというデータが入っています。
             */
            /*
             string filePath = download + "/mgs.csv";

             string strcsv = "";
             string strmail = "";
             foreach (KeyValuePair<string, int> pair in dicPC)
             {
                 //id,attach
                 strcsv = strcsv + (pair.Key + "," + pair.Value + "\n");
                 strmail = strmail + (pair.Key + "\t" + pair.Value + "\n");
             }
             foreach (KeyValuePair<string, int> pair in dicSP)
             {
                 //id,attach
                 strcsv = strcsv + (pair.Key + "," + pair.Value + "\n");
                 strmail = strmail + (pair.Key + "\t" + pair.Value + "\n");
             }
             string fileName = Path.GetFileNameWithoutExtension(filePath);
             FileStream fs = new FileStream("./" + fileName + ".csv", FileMode.Create);
             StreamWriter sw = new StreamWriter(fs);
             sw.WriteLine(strcsv);
             sw.Close();
             fs.Close();
             Mail("pa01m97ksuxa615@gmail.com", "xscBLjNyHmDb9yX", "pa01m97ksuxa615@gmail.com", "nameasakahikari@gmail.com", fileName + ".csv", strmail, "./" + fileName + ".csv");
             */
        }

        public void Fanza(IList<IList<Object>> values, IList<IList<Object>> valuesWrite)
        {

            DateTime targetdate = DateTime.Now.AddMonths(-int.Parse(ConfigurationManager.AppSettings.Get("FANZA_monthlate")));
            int year = targetdate.Year;
            int month = targetdate.Month;

            string userid = ConfigurationManager.AppSettings.Get("FANZA_userid");
            string password = ConfigurationManager.AppSettings.Get("FANZA_password");

            string download = System.Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "/Downloads";

            System.IO.File.Delete(download + "/商品別売上.csv");

            //DateTime date = new DateTime(year,month,1,0,0,0);
            string monthstr = month.ToString("D2");

            IWebDriver driver = new ChromeDriver();

            var wait = new WebDriverWait(driver, new TimeSpan(0, 0, 20));

            driver.Url = "https://owner-admin.dmm.com/owner";
            wait.Until(
                ExpectedConditions.ElementExists(By.Name("mail"))
            );
            IWebElement text = driver.FindElement(By.Name("mail"));
            text.SendKeys(userid);
            text = driver.FindElement(By.Name("password"));
            text.SendKeys(password);
            System.Threading.Thread.Sleep(2000);
            IWebElement btn = driver.FindElement(By.CssSelector("button[type='submit']"));

            btn.Click();


            wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("a[href='https://owner-admin.dmm.com/owner/search/product']")));
            btn = driver.FindElement(By.CssSelector("a[href='https://owner-admin.dmm.com/owner/search/product']"));
            btn.Click();

            wait.Until(
                ExpectedConditions.ElementExists(By.Name("dt_year_from"))
            );
            new SelectElement(driver.FindElement(By.Name("dt_year_from"))).SelectByValue(year.ToString());
            new SelectElement(driver.FindElement(By.Name("dt_month_from"))).SelectByValue(month.ToString());
            new SelectElement(driver.FindElement(By.Name("dt_day_from"))).SelectByValue("1");
            new SelectElement(driver.FindElement(By.Name("dt_year_to"))).SelectByValue(year.ToString());
            new SelectElement(driver.FindElement(By.Name("dt_month_to"))).SelectByValue(month.ToString());
            new SelectElement(driver.FindElement(By.Name("dt_day_to"))).SelectByValue(DateTime.DaysInMonth(year, month).ToString());

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
            /**
             * これを実行すると、sales.csvかattach.csvという名前で必要なファイルがダウンロードされると思います。
             */
            string filePath= download + "/商品別売上.csv";
            long filesize1, filesize2;
            do
            {
                Thread.Sleep(500);
                
            } while (!File.Exists(filePath));
            do
            {

                filesize1 = (new System.IO.FileInfo(filePath)).Length;  // check file size
                Thread.Sleep(500);      // wait for 5 seconds
                filesize2 = (new System.IO.FileInfo(filePath)).Length;  // check file size again

            } while (filesize1 != filesize2);


            StreamReader sr = new StreamReader(filePath, Encoding.GetEncoding("Shift_JIS"));

            //Console.WriteLine(sr.ReadToEnd());
            
            string date_ = (string)values[2][1];
            int column = 2;
            if (date_.Length >= 6)
            {
                int year_ = int.Parse(date_.Substring(0, 4));
                int month_ = int.Parse(date_.Substring(4, 2));
                column = (year - year_) * 12 + (month - month_) + 2;
                Console.WriteLine("date:" + date_ + " year:" + year_ + " month:" + month_ + " column:" + column);

            }
            addToList<IList<Object>>(valuesWrite, new List<Object>(), new List<Object>(), column);
            addToList<Object>(valuesWrite[column], year + "" + monthstr + "売上", null, 1);
            
            var dic = new Dictionary<String, int>();
            sr.ReadLine();
            for (string line=sr.ReadLine();line!=null;line=sr.ReadLine()){
                bool qu = false;
                char[] chararray = line.ToCharArray();
                for(int i=0;i< chararray.Length;i++){
                    char c = chararray[i];
                    switch (c)
                    {
                        case '"':
                            qu = !qu;
                            chararray[i] = ' ';
                            break;
                        case ',':
                            if (qu) chararray[i] = ' ';
                            break;
                    }
                }
                string[] elms = new string(chararray).Replace(" ", "").Split(",");
                //foreach (string e in elms) Console.WriteLine(e);
                Console.WriteLine(elms[6] + ","+elms[13]) ;
                int value = int.Parse(elms[13]);
                if (dic.ContainsKey(elms[6]))
                {
                    value += dic[elms[6]];
                    dic.Remove(elms[6]);
                    dic.Add(elms[6], value);
                }
                else
                {
                    dic.Add(elms[6], value);
                }
            }
            int index = 0;
            foreach (KeyValuePair<string, int> pair in dic)
            {
                index++;
                //id,attach
                string id = pair.Key;
                bool flag = false;
                for (int i = 0; i < values[0].Count; i++)
                {
                    //Console.WriteLine((string)values[0][i]);
                    if (((string)values[0][i]).Contains(id))
                    {
                        Console.WriteLine(pair.Key + " = " + values[0][i] + "(" + i + ")");
                        addToList<Object>(valuesWrite[column], pair.Value, "0", i);
                        flag = true;
                        break;
                    }
                }
                if (!flag)
                {
                    Console.WriteLine(pair.Key + " Not Found");
                }
            }
            }

        public void Mpo(IList<IList<Object>> values, IList<IList<Object>> valuesWrite)
        {

            int monthlate = int.Parse(ConfigurationManager.AppSettings.Get("mpo_monthlate"));
            DateTime targetdate = DateTime.Now.AddMonths(-monthlate);
            int year = targetdate.Year;
            int month = targetdate.Month;

            string userid = ConfigurationManager.AppSettings.Get("mpo_userid");
            string password = ConfigurationManager.AppSettings.Get("mpo_password");

            string download = System.Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "/Downloads";

            System.IO.File.Delete(download + "/sales.csv");

            //DateTime date = new DateTime(year,month,1,0,0,0);
            string monthstr = month.ToString("D2");

            IWebDriver driver = new ChromeDriver();

            var wait = new WebDriverWait(driver, new TimeSpan(0, 0, 20));

            driver.Url = @"http://mpo.jp/sp/provider/index.php";
            wait.Until(
                ExpectedConditions.ElementExists(By.Name("login_id"))
            );
            IWebElement text = driver.FindElement(By.Name("login_id"));
            text.SendKeys(userid);
            text = driver.FindElement(By.Name("login_pw"));
            text.SendKeys(password);
            System.Threading.Thread.Sleep(2000);
            IWebElement btn = driver.FindElement(By.CssSelector("input[type='submit']"));
            btn.Click();

            System.Threading.Thread.Sleep(2000);

            string sid = driver.Url.Substring(driver.Url.LastIndexOf("sid=")+4);

            driver.Url = @"http://mpo.jp/sp/provider/sales.php?sid="+sid;

            wait.Until(
                ExpectedConditions.ElementExists(By.CssSelector("input[type='submit']"))
            );

            btn = driver.FindElement(By.CssSelector("input[type='submit']"));

            btn.Click();

            /**
             * これを実行すると、sales.csvという名前で必要なファイルがダウンロードされると思います。
             */
            string filePath = download + "/sales.csv";
            long filesize1, filesize2;
            do
            {
                Thread.Sleep(500);

            } while (!File.Exists(filePath));
            do
            {

                filesize1 = (new System.IO.FileInfo(filePath)).Length;  // check file size
                Thread.Sleep(500);      // wait for 5 seconds
                filesize2 = (new System.IO.FileInfo(filePath)).Length;  // check file size again

            } while (filesize1 != filesize2);



            StreamReader sr = new StreamReader(filePath, Encoding.GetEncoding("Shift_JIS"));

            //Console.WriteLine(sr.ReadToEnd());

            string date_ = (string)values[2][1];
            int column = 2;
            if (date_.Length >= 6)
            {
                int year_ = int.Parse(date_.Substring(0, 4));
                int month_ = int.Parse(date_.Substring(4, 2));
                column = (year - year_) * 12 + (month - month_) + 2;
                Console.WriteLine("date:" + date_ + " year:" + year_ + " month:" + month_ + " column:" + column);

            }
            addToList<IList<Object>>(valuesWrite, new List<Object>(), new List<Object>(), column);
            addToList<Object>(valuesWrite[column], year + "" + monthstr + "売上", null, 1);

            var dic = new Dictionary<String, int>();
            sr.ReadLine();
            sr.ReadLine();
            for (string line = sr.ReadLine(); line != null; line = sr.ReadLine())
            {
                
                bool qu = false;
                char[] chararray = line.ToCharArray();
                for (int i = 0; i < chararray.Length; i++)
                {
                    char c = chararray[i];
                    switch (c)
                    {
                        case '"':
                            qu = !qu;
                            chararray[i] = ' ';
                            break;
                        case ',':
                            if (qu) chararray[i] = ' ';
                            break;
                    }
                }
                //Console.WriteLine(chararray);
                string[] elms = new string(chararray).Replace(" ", "").Split(",");
                if (elms.Length < 5 + monthlate) continue;
                //foreach (string e in elms) Console.WriteLine(e);
                Console.WriteLine(elms[1] + "," + elms[5+monthlate]);
                int value = int.Parse(elms[5 + monthlate]);
                if (dic.ContainsKey(elms[1]))
                {
                    value += dic[elms[1]];
                    dic.Remove(elms[1]);
                    dic.Add(elms[1], value);
                }
                else
                {
                    dic.Add(elms[1], value);
                }
            }
            int index = 0;
            foreach (KeyValuePair<string, int> pair in dic)
            {
                index++;
                //id,attach
                string id = pair.Key;
                bool flag = false;
                for (int i = 0; i < values[0].Count; i++)
                {
                    //Console.WriteLine((string)values[0][i]);
                    if (((string)values[0][i]).Contains(id.ToLower().Replace("-", "")))
                    {
                        Console.WriteLine(pair.Key + " = " + values[0][i] + "(" + i + ")");
                        addToList<Object>(valuesWrite[column], pair.Value, "0", i+1);
                        flag = true;
                        break;
                    }
                }
                if (!flag)
                {
                    Console.WriteLine(pair.Key + " Not Found");
                }
            }

            sr.Close();
            System.IO.File.Delete(download + "/sales.csv");


            driver.Url = @"http://mpo.jp/provider/index.php";
            wait.Until(
                ExpectedConditions.ElementExists(By.Name("login_id"))
            );
            text = driver.FindElement(By.Name("login_id"));
            text.SendKeys(userid);
            text = driver.FindElement(By.Name("login_pw"));
            text.SendKeys(password);
            System.Threading.Thread.Sleep(2000);
            btn = driver.FindElement(By.CssSelector("input[type='submit']"));
            btn.Click();

            System.Threading.Thread.Sleep(2000);

            sid = driver.Url.Substring(driver.Url.LastIndexOf("sid=") + 4);

            driver.Url = @"http://mpo.jp/provider/sales.php?sid=" + sid;

            wait.Until(
                ExpectedConditions.ElementExists(By.CssSelector("input[type='submit']"))
            );

            btn = driver.FindElement(By.CssSelector("input[type='submit']"));

            btn.Click();
            /**
             * これを実行すると、sales.csvという名前で必要なファイルがダウンロードされると思います。
             */
            filePath = download + "/sales.csv";
            do
            {
                Thread.Sleep(500);

            } while (!File.Exists(filePath));
            do
            {

                filesize1 = (new System.IO.FileInfo(filePath)).Length;  // check file size
                Thread.Sleep(500);      // wait for 5 seconds
                filesize2 = (new System.IO.FileInfo(filePath)).Length;  // check file size again

            } while (filesize1 != filesize2);

            sr = new StreamReader(filePath, Encoding.GetEncoding("Shift_JIS"));
            dic = new Dictionary<String, int>();
            sr.ReadLine();
            sr.ReadLine();
            for (string line = sr.ReadLine(); line != null; line = sr.ReadLine())
            {

                bool qu = false;
                char[] chararray = line.ToCharArray();
                for (int i = 0; i < chararray.Length; i++)
                {
                    char c = chararray[i];
                    switch (c)
                    {
                        case '"':
                            qu = !qu;
                            chararray[i] = ' ';
                            break;
                        case ',':
                            if (qu) chararray[i] = ' ';
                            break;
                    }
                }
                //Console.WriteLine(chararray);
                string[] elms = new string(chararray).Replace(" ", "").Split(",");
                if (elms.Length < 5 + monthlate) continue;
                //foreach (string e in elms) Console.WriteLine(e);
                Console.WriteLine(elms[1] + "," + elms[5 + monthlate]);
                int value = int.Parse(elms[5 + monthlate]);
                if (dic.ContainsKey(elms[1]))
                {
                    value += dic[elms[1]];
                    dic.Remove(elms[1]);
                    dic.Add(elms[1], value);
                }
                else
                {
                    dic.Add(elms[1], value);
                }
            }
            index = 0;
            foreach (KeyValuePair<string, int> pair in dic)
            {
                index++;
                //id,attach
                string id = pair.Key;
                bool flag = false;
                for (int i = 0; i < values[0].Count; i++)
                {
                    //Console.WriteLine((string)values[0][i]);
                    if (((string)values[0][i]).Contains(id.ToLower().Replace("-", "")))
                    {
                        Console.WriteLine(pair.Key + " = " + values[0][i] + "(" + i + ")");
                        addToList<Object>(valuesWrite[column], pair.Value, "0", i);
                        flag = true;
                        break;
                    }
                }
                if (!flag)
                {
                    Console.WriteLine(pair.Key + " Not Found");
                }
            }
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

