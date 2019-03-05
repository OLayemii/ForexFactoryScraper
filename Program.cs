using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Threading.Tasks;
using HtmlAgilityPack;
using System.Web;
using System.Text.RegularExpressions;
using CsvHelper;
using System.IO;

namespace ForexCalendarScraper
{

    internal class RowRecord
    {
        public string Date { get; set; }
        public string Time { get; set; }
        public string Currency { get; set; }
        public string Impact { get; set; }
        public string Name { get; set; }
        public string Detail { get; set; }
        public string Actual { get; set; }
        public string Forecast { get; set; }
        public string Previous { get; set; }
        public string Graph { get; set; }
    }
    class FFScraper
    {
        private static List<RowRecord> records = new List<RowRecord>();
        private static HtmlDocument html;
        static void Main(string[] args)
        {

            try
            {
                
                Console.WriteLine("Enter Begin Date For Scraping (YYYY-MM-DD)");
                String beginDateStr = Console.ReadLine();
                string[] beginDateParts = beginDateStr.Split('-');

                DateTime beginDate = new DateTime(Convert.ToInt32(beginDateParts[0]), Convert.ToInt32(beginDateParts[1]), Convert.ToInt32(beginDateParts[2]));


                Console.WriteLine("Enter End Date For Scraping (YYYY-MM-DD)");
                String endDateStr = Console.ReadLine();
                string[] endDateParts = endDateStr.Split('-');

                DateTime endDate = new DateTime(Convert.ToInt32(endDateParts[0]), Convert.ToInt32(endDateParts[1]), Convert.ToInt32(endDateParts[2]));

                int dateOffset = Convert.ToInt32((endDate - beginDate).TotalDays);
                string dateStringForScrape;
                if (dateOffset < 0) {
                    Console.WriteLine("Date Elapsed");
                }
                else
                {
                    Console.WriteLine("Scraper  Started. . .");
                    for (int i = 0; i <= dateOffset; i++)
                    {
                        var newDay = beginDate.AddDays(i);
                        dateStringForScrape = newDay.ToString("MMM")+newDay.Day+"."+newDay.Year;
                        dateStringForScrape = dateStringForScrape.ToLower();

                        
                        string url = "https://www.forexfactory.com/calendar.php?day="+dateStringForScrape;
                        scrapePage(url);


                        makeCSV(records, newDay.ToString("MMMM")+"_"+newDay.Day+"_"+newDay.Year);
                    }

                }
            }
            catch (ArgumentOutOfRangeException e)
            {
                Console.WriteLine("Please revise your date passed: "+e.Message);
            }

            Console.ReadLine();
        }

        public static void scrapePage(string url)
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3 | SecurityProtocolType.Tls | SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;


                WebClient webpage = new WebClient();
                webpage.Headers.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/72.0.3626.119 Safari/537.36");
                string source = webpage.DownloadString(url);

                html = new HtmlDocument();

                html.LoadHtml(source);

//              Console.WriteLine(getTableHeader());
                getTableRows();

                return;

            }
            catch (WebException e)
            {
                Console.WriteLine(e.Message);
            }

            Console.ReadLine();
        }


        public static List<dynamic> getTableRows()
        {
            var dataRows = html.DocumentNode.SelectNodes("//tr[contains(@class, 'calendar_row') and contains(@class, 'calendar__row')]");
            var header = new List<String>{"Date", "Time", "Currency", "Impact", "Name", "Detail", "Actual", "Forecast", "Previous", "Graph"};
            try
            {
                foreach (HtmlNode dRows in dataRows)
                {

                    int counter = 0;
                    var record = new RowRecord();
                    foreach (HtmlNode dColumns in dRows.SelectNodes("td"))
                    {
                        var str = "";
                        if (header[counter].ToLower() == "impact"){
                            var nodeTitle = dColumns.SelectNodes("div/span");
                            str = nodeTitle[0].Attributes["title"].Value;
                        }
                        else
                        {
                            str = dColumns.InnerText.Trim('\r', '\n');
                        }
                        record.GetType().GetProperty(header[counter]).SetValue(record, str);
                        counter++;
                    }
                    records.Add(record);
                }

            }catch(Exception e){
                Console.WriteLine(e.Message);
            }

            return null;
        }

        public static HtmlNode removeBlanks(HtmlNode removedSpaces)
        {
            return removedSpaces;
        }

        public static List<string> getTableHeader()
        {
            try
            {
                List<string> tableHeader = new List<string>();
                var dataTable = html.DocumentNode.SelectNodes("//table[@class='calendar__table']/thead/tr[contains(@class, 'calendar__header--desktop') and contains(@class, 'subhead')]/th");
                foreach (HtmlNode tableColumn in dataTable)
                {
                    var tableColumnText = removeHtmlAttr(tableColumn.InnerText);
                    tableHeader.Add(tableColumnText);

                }
                return tableHeader;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }

            return null;

        }


        public static String removeHtmlAttr(string str)
        {
            string text = Regex.Replace(str, @"<[^>]+>|&nbsp;", "").Trim();
            text = Regex.Replace(text, @"\s{2,}", " ");

            return text;
        }

        public static void makeCSV(List<RowRecord> csvdata, String filename)
        {
            var documentDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            var dirName = System.IO.Path.Combine(documentDir, "ForexFactoryCSV");
            var path = System.IO.Path.Combine(dirName, filename+".csv");
            if (!Directory.Exists(dirName))
            {
                Directory.CreateDirectory(dirName);
            }

            using (var writer = new StreamWriter(path))
            using (var csv = new CsvWriter(writer))
            {
                csv.WriteRecords(csvdata);
            }

            Console.WriteLine("Saved Scraped CSV At "+path);
            records.Clear();
            return;
        }
    }

}
