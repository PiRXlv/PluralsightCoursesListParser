using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using HtmlAgilityPack;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace PluralsightCoursesListParser
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            const string Url = @"http://pluralsight.com/training/Courses/CoursesContent";
            const string FileName = @"Courses list.xlsx";

            var doc = GetHtml(Url);

            var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("Courses list");

            WriteTableHeader(sheet);
            ProcessHtml(doc, sheet);
            WriteWorkbook(FileName, workbook);

            Process.Start(FileName);
        }

        private static void WriteWorkbook(string FileName, XSSFWorkbook workbook)
        {
            using (var stream = new FileStream(FileName, FileMode.Create))
            {
                workbook.Write(stream);
            }
        }

        private static void ProcessHtml(HtmlDocument doc, ISheet sheet)
        {
            HtmlNodeCollection headers = doc.DocumentNode.SelectNodes("//div[contains(@class, 'categoryHeader')]");
            foreach (var headerNode in headers)
            {
                ProcessHeaderNode(headerNode, sheet);
            }
        }

        private static HtmlDocument GetHtml(string Url)
        {
            var web = new HtmlWeb();
            HtmlDocument doc = web.Load(Url);
            return doc;
        }

        private static void WriteTableHeader(ISheet sheet)
        {
            var columnHeaders = new[]
                                    {
                                        "Category", "Title", "Description", "Level", "Rating", "Duration", "Authors", 
                                        "Released", "Url"
                                    };
            IRow row = sheet.CreateRow(0);
            for (int i = 0; i < columnHeaders.Length; i++)
            {
                ICell cell = row.CreateCell(i);
                cell.SetCellValue(columnHeaders[i]);
            }
        }

        private static void ProcessCourseNodes(HtmlNode courseNode, IRow row)
        {
            var values = ExtractRecordFields(courseNode);
            
            for (int i = 0; i < values.Length; i++)
            {
                ICell cell = row.CreateCell(i + 1);
                cell.SetCellValue(values[i]);
                if (i == 1)
                {
                    cell.CellStyle.WrapText = true;
                }
            }
        }

        private static string[] ExtractRecordFields(HtmlNode courseNode)
        {
            string title = string.Empty;
            string description = string.Empty;
            string level = string.Empty;
            string rating = string.Empty;
            string duration = string.Empty;
            string released = string.Empty;
            string authors = string.Empty;
            string url = string.Empty;

            foreach (var node in courseNode.ChildNodes)
            {
                if (node.Attributes["class"] == null)
                {
                    continue;
                }

                string nodeClass = node.Attributes["class"].Value;
                switch (nodeClass)
                {
                    case "title":
                        title = node.FirstChild.NextSibling.InnerText;
                        description = node.FirstChild.NextSibling.Attributes["title"].Value;
                        url = "http://pluralsight.com" + node.FirstChild.NextSibling.Attributes["href"].Value;
                        break;
                    case "rating":
                        rating =
                            node.FirstChild.NextSibling.Attributes["title"].Value.Split(
                                new[] { " " },
                                StringSplitOptions.RemoveEmptyEntries).First();
                        break;
                    case "level":
                        level = node.FirstChild.InnerText.Trim();
                        break;
                    case "duration":
                        duration = node.FirstChild.InnerText.Trim().Replace("[", string.Empty).Replace("]", string.Empty);
                        break;
                    case "releaseDate":
                        released = node.FirstChild.InnerText.Trim();
                        break;
                    case "author":
                        authors = string.Join(", ", node.SelectNodes(".//a").Select(n => n.InnerText.Trim()));
                        break;
                }
            }

            var values = new[] { title, description, level, rating, duration, authors, released, url };
            return values;
        }

        private static void ProcessHeaderNode(HtmlNode headerNode, ISheet sheet)
        {
            HtmlNode titleNode = headerNode.SelectSingleNode(".//div[contains(@class, 'title')]");
            string category = titleNode.InnerText.Trim();
            Console.WriteLine(category);
            HtmlNodeCollection courses = headerNode.NextSibling.NextSibling.SelectNodes(".//tr");
            int rowNum = sheet.LastRowNum + 1;
            foreach (var courseNode in courses)
            {
                IRow row = sheet.CreateRow(rowNum);
                row.CreateCell(0).SetCellValue(category);
                ProcessCourseNodes(courseNode, row);
                rowNum++;
            }
        }
    }
}