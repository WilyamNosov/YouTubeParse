using OfficeOpenXml;
using OfficeOpenXml.Style;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;

namespace YoutubeParse
{
    class Program
    {
        public string[] dateFB = DateTime.Now.ToString().Split('.');
        public ExcelPackage excel = new ExcelPackage();

        public IWebDriver driver = new ChromeDriver(Environment.CurrentDirectory + Path.DirectorySeparatorChar + "");

        //public IWebDriver driver = new ChromeDriver(@".InputData\\chromedriver_win32");
        public string parseUrl = "";

        public List<string> urlList = new List<string>();
        public List<string> namesList = new List<string>();

        public List<string> ExcelNames = new List<string>();
        public List<string> ExcelUrls = new List<string>();
        public List<List<string>> resultList = new List<List<string>>();

        public List<string> zeroUserUrlList = new List<string>();
        public List<string> zeroUserNameList = new List<string>();

        public void scrollToBottom()
        {
            List<IWebElement> toBottom = driver.FindElements(By.TagName("ytd-grid-channel-renderer")).ToList();
            Actions actions = new Actions(driver);
            int subscribersCount = 0;
            if (toBottom.Count != 0)
            {
                while (true)
                {
                    System.Threading.Thread.Sleep(2000);
                    toBottom = driver.FindElements(By.TagName("ytd-grid-channel-renderer")).ToList();
                    actions.MoveToElement(toBottom.ElementAt(toBottom.Count - 1));
                    actions.Perform();
                    if (subscribersCount == toBottom.Count)
                    {
                        break;
                    }
                    subscribersCount = toBottom.Count;
                }
            }
        }

        public void getInfo()
        {
            ReadOnlyCollection<IWebElement> chanelUrls = driver.FindElements(By.Id("channel-info"));
            foreach (IWebElement url in chanelUrls)
            {
                urlList.Add(url.GetAttribute("href"));
                namesList.Add(url.FindElement(By.TagName("span")).Text);
            }
        }

        public void getUrlsFromExcel()
        {
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(Environment.CurrentDirectory + Path.DirectorySeparatorChar + "InputData\\excl.xlsx")))
            {
                ExcelWorksheet myWorksheet = xlPackage.Workbook.Worksheets.First(); //select sheet here
                int totalRows = myWorksheet.Dimension.End.Row;

                for (int i = 1; i < totalRows; i++)
                {
                    ExcelNames.Add(myWorksheet.GetValue(i, 1).ToString());
                    ExcelUrls.Add(myWorksheet.GetValue(i, 2).ToString());
                }
            }
        }

        public void buildResultsList()
        {
            this.getUrlsFromExcel();
            for (int i = 0; i < ExcelUrls.Count; i++)//ExcelUrls.Count; i++)
            {
                string url = ExcelUrls.ElementAt(i);
                this.setDriver(url);
                this.scrollToBottom();
                this.getInfo();
                for (int j = 0; j < urlList.Count; j++)
                {
                    resultList.Add(new List<string>() { urlList.ElementAt(j), namesList.ElementAt(j) });
                }
                this.createExcel(ExcelNames.ElementAt(i), i, url);
                urlList.Clear();
                namesList.Clear();
                resultList.Clear();
            }
            this.addZeroUsers();
            this.saveToExcel();
        }

        public void createExcel(string sheetName, int index, string url)
        {
            if (resultList.Count > 0)
            {
                ExcelWorksheet workSheet = excel.Workbook.Worksheets.Add(sheetName + " " + index);
                workSheet.TabColor = System.Drawing.Color.Black;
                for (int i = 0; i < urlList.Count; i++)
                {
                    for (int j = 0; j < resultList.Count; j++)
                    {
                        workSheet.Cells[j + 1, 1].Value = resultList.ElementAt(j).ElementAt(0);
                        workSheet.Cells[j + 1, 2].Value = resultList.ElementAt(j).ElementAt(1);
                    }
                }
            }
            else
            {
                zeroUserUrlList.Add(url);
                zeroUserNameList.Add(sheetName);
            }

        }

        public void addZeroUsers()
        {
            ExcelWorksheet workSheet = excel.Workbook.Worksheets.Add("ZeroUsers");
            workSheet.TabColor = System.Drawing.Color.Black;
            for (int i = 0; i < zeroUserNameList.Count; i++)
            {
                workSheet.Cells[i + 1, 1].Value = zeroUserUrlList.ElementAt(i);
                workSheet.Cells[i + 1, 2].Value = zeroUserNameList.ElementAt(i);
            }

        }

        public void saveToExcel()
        {
            string s = DateTime.Now.ToString("dd_MMMM_yyyy HH_mm_ss");
            Console.WriteLine(s);

            string filePath = Environment.CurrentDirectory + Path.DirectorySeparatorChar + "OutputData\\ExcelParse" + s + ".xlsx";

            FileInfo fi = new FileInfo(filePath);
            excel.SaveAs(fi);
        }

        public void setDriver(string url)
        {
            driver.Url = (url + "/channels");
        }

        public void quit()
        {
            driver.Quit();
        }

        static void Main(string[] args)
        {
            Program A = new Program();
            A.buildResultsList();
            A.quit();
        }
    }
}
