using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;

namespace Spider01
{
    class Program
    {
        static async Task Main(string[] args)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("tretta");
                worksheet.Column(1).SetDataType(XLDataType.Text);
                int index = 1;
                for (int i = 794; i <= 893; i++)
                {
                    var url = $@"https://www.pokemontretta.com.tw/trettaBox.php?id={i}&ajax=true&height=435";

                    HtmlWeb web = new HtmlWeb();

                    var htmlDoc = await web.LoadFromWebAsync(url);
                    while (htmlDoc.Text.Contains("PROXY_AUTH_REQUIRED"))
                    {
                        htmlDoc = await web.LoadFromWebAsync(url);
                        Thread.Sleep(1000);
                    }

                    var ths = htmlDoc.DocumentNode.SelectNodes("//div/div/table/tr/th");
                    var tds = htmlDoc.DocumentNode.SelectNodes("//div/div/table/tr/td");

                    Console.WriteLine(ths[0].InnerText);


                    if (!string.IsNullOrWhiteSpace(ths[0].InnerText))
                    {
                        var energy = tds[2].InnerText;
                        if (tds[2].InnerText.Contains("("))
                        {
                            energy = tds[2].InnerText.Substring(0, tds[2].InnerText.IndexOf('('));
                        }

                        var attack = ths.First(x => x.InnerText == "攻擊").NextSibling.NextSibling.InnerText;
                        var defense = ths.First(x => x.InnerText == "防禦").NextSibling.NextSibling.InnerText;
                        var speed = ths.First(x => x.InnerText == "速度").NextSibling.NextSibling.InnerText;

                        worksheet.Cell(index, 1).Value = ths[0].InnerText + " sss";
                        worksheet.Cell(index, 2).Value = ths[1].InnerText;
                        worksheet.Cell(index, 3).Value = @""", Attributes: """;
                        worksheet.Cell(index, 4).Value = tds[1].InnerText;
                        worksheet.Cell(index, 5).Value = @""", Energy: """;
                        worksheet.Cell(index, 6).Value = energy;
                        worksheet.Cell(index, 7).Value = @""", Life: """;
                        worksheet.Cell(index, 8).Value = tds[3].InnerText;
                        worksheet.Cell(index, 9).Value = @""", Attack: """;
                        worksheet.Cell(index, 10).Value = attack;
                        worksheet.Cell(index, 11).Value = @""", Defense: """;
                        worksheet.Cell(index, 12).Value = defense;
                        worksheet.Cell(index, 13).Value = @""", Speed: """;
                        worksheet.Cell(index, 14).Value = speed;
                        worksheet.Cell(index, 15).Value = @"""}, ";

                        HtmlNode imgNode = htmlDoc.DocumentNode.SelectSingleNode("//*[@id=\"mainImg01\"]");
                        var byteFile = await DownloadFile(imgNode.Attributes.First(x => x.Name == "src").Value);
                        if (byteFile != null)
                        {
                            Directory.CreateDirectory("img");
                            File.WriteAllBytes($"img/{ths[0].InnerText}.gif", byteFile);
                        }
                    }
                    index++;
                }
                workbook.SaveAs("tretta.xlsx");
            }
            Console.WriteLine("Done");
            Console.ReadKey();
        }

        public static async Task<byte[]> DownloadFile(string url)
        {
            using (var client = new HttpClient())
            {

                using (var result = await client.GetAsync(url))
                {
                    if (result.IsSuccessStatusCode)
                    {
                        return await result.Content.ReadAsByteArrayAsync();
                    }

                }
            }
            return null;
        }
    }
}
