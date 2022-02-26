using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
//using Aspose.Cells;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using wbExample.Models;
using Spire.Xls;
using System.Text.RegularExpressions;
using HtmlAgilityPack;

namespace wbExample.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class LastDanceController : ControllerBase
    {
        DateTime selectedDate;
        public LastDanceController()
        {

            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ru-RU");
        }
        [HttpGet("getnew")]
        public ActionResult<IEnumerable<string>> GetNew() //вернуть информацию о новом расписании
        {
            Differents.DateOut(DateTime.Now.AddDays(7));
            try
            {
                string trim1 = Differents.upMonth.Substring(0, 3);
                var workbook = new Aspose.Cells.Workbook(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/_{Differents.downDay}_{Differents.upDay}_{trim1}.xlsx");
                Differents.DateOut(DateTime.Now);
            }
            catch
            {
                if (newRaspisanie() == "bad")
                    return new ObjectResult("нет нового расписания");
            }
            return new ObjectResult("есть новое расписание");
        }
        [HttpGet("getteacher/{teacher}")]
        public ActionResult<IEnumerable<List<DayWeekClass>>> GetTeach(string teacher)
        {
            List<DayWeekClass> days = new List<DayWeekClass>();
            for (int i = 1; i <= 6; i++)
            {
                int row = 6;
                List<DayWeekClass> teachers = Differents.raspisanieteach(row * i, teacher);
                days.AddRange(teachers.ToArray());

            }
            return new ObjectResult(days);
        }


        [HttpGet("getgroup/{group}")]
        public ActionResult<IEnumerable<List<DayWeekClass>>> Get(string group) //вернуть расписание по группам
        {
            List<DayWeekClass> days = new List<DayWeekClass>();
            int column = Differents.IndexGroup(group);
            for (int j = 1; j <= 6; j++)
            {
                List<DayWeekClass> metrics = Differents.EnumerableMetrics(j * 6, column);
                days.AddRange(metrics.ToArray());
            }
            return new ObjectResult(days);
        }

        [HttpGet("getcabinet/{kabinet}")] //вернуть расписание по кабинетам
        public ActionResult<IEnumerable<List<DayWeekClass>>> GetKab(string kabinet)
        {
            List<DayWeekClass> days = new List<DayWeekClass>();
            for (int i = 1; i <= 6; i++)
            {
                int row = 6;
                List<DayWeekClass> kabinets = Differents.raspisaniekab(row * i, kabinet);
                days.AddRange(kabinets.ToArray());

            }
            return new ObjectResult(days);
        }
        [HttpGet("getversionMobile")] //вернуть расписание по кабинетам
        public ActionResult<string> GetVarsion()
        {
            using (StreamReader sr = new StreamReader($"{AppDomain.CurrentDomain.BaseDirectory}Version/mobileVersion.txt"))
            {
                return new ObjectResult(sr.ReadToEnd());
            }
        }
        [HttpGet("getversionDesktop")] //вернуть расписание по кабинетам
        public ActionResult<string> GetVersion()
        {
            using (StreamReader sr = new StreamReader($"{AppDomain.CurrentDomain.BaseDirectory}Version/desktopVersion.txt"))
            {
                return new ObjectResult(sr.ReadToEnd());
            }
        }

        [HttpGet("getdate/{date}")] //вернуть расписание по конкретной неделе
        public ActionResult<IEnumerable<List<DayWeekClass>>> GetDate(string date)
        {
            DateTime time = DateTime.Parse(date);
            selectedDate = time;
            Differents.DateOut(time);
            string trim1 = Differents.upMonth.Substring(0, 3);
            if (System.IO.File.Exists(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/_{Differents.downDay}_{Differents.upDay}_{trim1}.xlsx"))
            {
                Differents._workbook = new XLWorkbook(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/_{Differents.downDay}_{Differents.upDay}_{trim1}.xlsx");
                HtmlDocument doc = new HtmlDocument();
                var web1 = new HtmlWeb();
                doc = web1.Load("https://oksei.ru/studentu/raspisanie_uchebnykh_zanyatij");
                var node = doc.DocumentNode.SelectSingleNode("//*[@class='container bg-white p-25 box-shadow-right radius']/p/a");
                var href = node.Attributes["href"].Value;
                var value = node.InnerText;
                using (var stream = new StreamWriter($"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/Formats.txt", false))
                {
                    stream.Write(value);
                    stream.Close();
                }
            }
            else
            {
                FormatsRaspisanie();
            }
            Differents.workSheet = Differents._workbook.Worksheets.First();
            List<DayWeekClass> days = new List<DayWeekClass>();
            Differents.DownloadFeatures(time);
            return new ObjectResult(days);
        }

        // GET api/users/5
        [HttpGet]
        [Route("getcabinetsList")]
        public async Task<ActionResult<IEnumerable<List<string>>>> Get() //вернуть список кабиентов
        {
            bool exit = true;
            List<string> dataforcb = new List<string>();
            for (int i = 3; i <= Differents.workSheet.ColumnsUsed().Count(); i++)
            {
                for (int j = 6; j <= Differents.workSheet.RowsUsed().Count(); j++)
                {
                    string result = Differents.workSheet.Cell(j, i).GetValue<string>();
                    if (result != "" && result != " " && result.Length > 3)
                    {
                        result = result.Remove(0, result.Length - 3).Trim();
                        if (result != "-" && result != "" && result.Length > 1)
                        {
                            Regex regex = new Regex("[0-9]{2,3}");
                            if (regex.IsMatch(result))
                            {
                                foreach (string output in dataforcb)
                                {
                                    if (output == result)
                                    {
                                        exit = false;
                                        break;
                                    }
                                }
                                if (exit)
                                {
                                    dataforcb.Add(result);
                                }
                                exit = true;
                            }
                        }
                    }
                }
            }
            dataforcb.Sort();
            return new ObjectResult(dataforcb);
        }
        [HttpGet]
        [Route("getteachersList")]
        public async Task<ActionResult<IEnumerable<List<string>>>> GetTeachList() //вернуть список кабиентов
        {
            bool exit = true;
            List<string> dataforcb = new List<string>();
            for (int i = 3; i <= Differents.workSheet.ColumnsUsed().Count(); i++)
            {
                for (int j = 6; j <= Differents.workSheet.RowsUsed().Count(); j++)
                {
                    string result = Differents.workSheet.Cell(j, i).GetValue<string>();
                    if (result != "" && result != " " && result.Length > 3)
                    {
                        if (result.Contains("ДОП"))
                        {
                            string[] massiv = result.Split(new char[] { '(', ')' });
                            result = massiv[1].Trim();
                        }
                        else
                        {
                            try
                            {
                                if (result.Length == 4)
                                {
                                    continue;
                                }
                                string[] massiv = result.Split('\n');
                                if (massiv.Length == 1)
                                { continue; }
                                result = massiv[1].Trim();
                            }
                            catch
                            { continue; }
                        }
                        if (result != "-" && result != "")
                        {
                            Regex regex = new Regex(@"[а-яА-Я]+\s[А-Я]{1}\.[А-Я]{1}\.?$");
                            if (regex.IsMatch(result))
                            {
                                foreach (string output in dataforcb)
                                {
                                    if (output == result)
                                    {
                                        exit = false;
                                        break;
                                    }
                                }
                                if (exit)
                                {
                                    dataforcb.Add(result);
                                }
                                exit = true;
                            }
                        }
                    }
                }
            }
            dataforcb.Sort();
            return new ObjectResult(dataforcb);

        }

        [HttpGet]
        [Route("getgroupList")]
        public async Task<ActionResult<IEnumerable<List<string>>>> get() //вернуть список групп
        {
            List<string> dataforcb = new List<string>();
            for (int i = 3; i <= Differents.workSheet.ColumnsUsed().Count(); i++)
            {
                dataforcb.Add(Differents.workSheet.Cell(5, i).GetValue<string>());
            }
            dataforcb.Sort();
            return new ObjectResult(dataforcb);
        }

        public static void RaspisanieIzm(XLWorkbook _workbook1, int h) //outcbKabinet
        {
            DateTime dateIZM = DateTime.Today;
            int dayWeek = (int)DateTime.Today.DayOfWeek;
            var worksheet = _workbook1.Worksheets.First();
            for (int i = 1; i <= worksheet.ColumnsUsed().Count(); i++)
            {
                int n = worksheet.RowsUsed().Count();
                for (int j = 11; j <= worksheet.RowsUsed().Count() + 10; j++)
                {
                    for (int l = 3; l <= Differents.workSheet.ColumnsUsed().Count(); l++)
                    {
                        if (Differents.workSheet.Cell(5, l).GetValue<string>() == worksheet.Cell(j, i).GetValue<string>())
                        {
                            bool a = false;
                            bool b = true;
                            bool c = false;
                            int g = 6;
                            for (int m = 1; m <= g; m++)
                            {
                                if (h <= 4)
                                {
                                    if (h == 4 && b)
                                    {
                                        g++;
                                        b = false;
                                    }
                                    IXLCell leg = worksheet.Cell(j + m, i);
                                    if (leg.Style.Font.FontSize >= 22 || leg.Value.ToString() == "" || a || leg.Value.ToString().Length == 4)
                                    {
                                        if (m >= 4 && h == 4)
                                        {
                                            if (Differents.workSheet.Cell(27, 2).Value.ToString() != "4")
                                                Differents.workSheet.Cell((6 * h) + m, l).Value = " ";
                                            else
                                                Differents.workSheet.Cell((6 * h) + m - 1, l).Value = " ";
                                        }
                                        else
                                        {
                                            Differents.workSheet.Cell((6 * h) + m - 1, l).Value = " ";
                                        }
                                        a = true;
                                    }
                                    else
                                    {
                                        if (m >= 4 && h == 4)
                                        {
                                            if (worksheet.Cell(j + m, i).Value.ToString().Contains("ЧКР") || c)
                                            {
                                                Differents.workSheet.Cell((6 * h) + m - 1, l).Value = worksheet.Cell(j + m, i);
                                                c = true;
                                            }
                                            else
                                            {
                                                if (Differents.workSheet.Cell(27, 2).Value.ToString() != "4")
                                                    Differents.workSheet.Cell((6 * h) + m, l).Value = worksheet.Cell(j + m, i);
                                                else
                                                    Differents.workSheet.Cell((6 * h) + m - 1, l).Value = worksheet.Cell(j + m, i);
                                            }

                                        }
                                        else
                                            Differents.workSheet.Cell((6 * h) + m - 1, l).Value = worksheet.Cell(j + m, i);
                                    }
                                }
                                else
                                {
                                    IXLCell leg = worksheet.Cell(j + m, i);
                                    if (leg.Style.Font.FontSize >= 22 || leg.Value.ToString() == "" || a || leg.Value.ToString().Length == 4)
                                    {
                                        if (Differents.workSheet.Cell(27, 2).Value.ToString() != "4")
                                            Differents.workSheet.Cell((6 * h) + m, l).Value = " ";
                                        else
                                            Differents.workSheet.Cell((6 * h) + m - 1, l).Value = " ";
                                        a = true;
                                    }
                                    else
                                    {
                                        if (Differents.workSheet.Cell(27, 2).Value.ToString() != "4")
                                            Differents.workSheet.Cell((6 * h) + m, l).Value = worksheet.Cell(j + m, i);
                                        else
                                            Differents.workSheet.Cell((6 * h) + m - 1, l).Value = worksheet.Cell(j + m, i);
                                    }
                                }
                            }
                        }
                    }
                }
            }


        }

        [HttpGet]
        [Route("getnewWeek/{btnContent}")]
        public async Task<ActionResult<DayWeekClass>> GetNewn(string btnContent)
        {
            if (btnContent == "Новое расписание!!!")
            {
                Differents.DateOut(DateTime.Now.AddDays(7));
                Differents._workbook = new XLWorkbook(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/_{Differents.downDay}_{Differents.upDay}_{Differents.upMonth.Substring(0, 3)}.xlsx");
            }
            else
            {
                Differents.DateOut(DateTime.Now);
                Differents._workbook = new XLWorkbook($"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/_{Differents.downDay}_{Differents.upDay}_{Differents.upMonth.Substring(0, 3)}.xlsx");
            }
            Differents.workSheet = Differents._workbook.Worksheets.First();
            DayWeekClass dk = new DayWeekClass();
            return Ok(dk);
        }
        [HttpGet]
        [Route("getSignData/{data}")]
        public async Task<ActionResult> GetSign(string data)
        {
            if (data == "Mat'NeTrogai")
                return Ok();
            else
                return BadRequest();
        }

        private string FormatsRaspisanie()
        {
            HtmlDocument doc = new HtmlDocument();
            var web1 = new HtmlWeb();
            doc = web1.Load("https://oksei.ru/studentu/raspisanie_uchebnykh_zanyatij");
            var node = doc.DocumentNode.SelectSingleNode("//*[@class='container bg-white p-25 box-shadow-right radius']/p/a");
            var href = node.Attributes["href"].Value;
            var value = node.InnerText;
            WebClient web = new WebClient();
            web.DownloadFile($"https://oksei.ru{href}", $"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/_{Differents.DdownDay.Day}_{Differents.DupDay.Day}_{Differents.upMonth.Substring(0, 3)}.xls");
            var workbook = new Aspose.Cells.Workbook(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/_{Differents.downDay}_{Differents.upDay}_{Differents.upMonth.Substring(0, 3)}.xls");
            workbook.Save(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/_{Differents.downDay}_{Differents.upDay}_{Differents.upMonth.Substring(0, 3)}.xlsx", Aspose.Cells.SaveFormat.Xlsx);
            Differents._workbook = new XLWorkbook(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/_{Differents.downDay}_{Differents.upDay}_{Differents.upMonth.Substring(0, 3)}.xlsx");
            using (var stream = new StreamWriter($"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/Formats.txt", false))
            {
                stream.Write(value);
                stream.Close();
            }
            return "good";
        }

        private string newRaspisanie()
        {
            HtmlDocument doc = new HtmlDocument();
            var web1 = new HtmlWeb();
            doc = web1.Load("https://oksei.ru/studentu/raspisanie_uchebnykh_zanyatij");
            var node = doc.DocumentNode.SelectSingleNode("//*[@class='container bg-white p-25 box-shadow-right radius']/p/a");
            var href = node.Attributes["href"].Value;
            var value = node.InnerText;
            using (var stream = new StreamReader($"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/Formats.txt"))
            {
                string l = stream.ReadToEnd();
                if (l != value)
                {
                    WebClient web = new WebClient();
                    web.DownloadFile($"https://oksei.ru{href}", $"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/_{Differents.DdownDay.Day}_{Differents.DupDay.Day}_{Differents.upMonth.Substring(0, 3)}.xls");
                    var workbook = new Aspose.Cells.Workbook(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/_{Differents.downDay}_{Differents.upDay}_{Differents.upMonth.Substring(0, 3)}.xls");
                    workbook.Save(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/_{Differents.downDay}_{Differents.upDay}_{Differents.upMonth.Substring(0, 3)}.xlsx", Aspose.Cells.SaveFormat.Xlsx);
                    Differents._workbook = new XLWorkbook(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/_{Differents.downDay}_{Differents.upDay}_{Differents.upMonth.Substring(0, 3)}.xlsx");
                    Differents.DateOut(DateTime.Now);
                    return "good";
                }
                else
                {
                    Differents.DateOut(DateTime.Now);
                    return "bad";
                }
            }
        }



        [HttpGet]
        [Route("addFormat/{format}")]
        public async Task<ActionResult> AddNewFormat(string format)
        {
            using (var writer = new StreamWriter($"{AppDomain.CurrentDomain.BaseDirectory}/Raspisanie/Formats.txt", true, System.Text.Encoding.Default))
            {
                writer.WriteLine(format);
                writer.Close();
            }
            return Ok();
        }

        [HttpGet]
        [Route("DeleteIzm/{delete}")]
        public async Task<ActionResult> DeleteIzm(string delete)
        {
            using (var reader = new StreamReader($"{AppDomain.CurrentDomain.BaseDirectory}/Raspisanie/allizm.txt"))
            {
                string s = reader.ReadToEnd();
                s = s.Replace(delete + ".xlsx", "");
                reader.Close();
                using var writer = new StreamWriter($"{AppDomain.CurrentDomain.BaseDirectory}/Raspisanie/allizm.txt", false, System.Text.Encoding.Default);
                writer.WriteLine(s);
                writer.Close();
                System.IO.File.Delete($"{AppDomain.CurrentDomain.BaseDirectory}/Raspisanie/{delete}.xlsx");
                System.IO.File.Delete($"{AppDomain.CurrentDomain.BaseDirectory}/Raspisanie/{delete}.xls");
            }
            return Ok();
        }

        [HttpGet]
        [Route("allIzm")]
        public async Task<ActionResult<string>> ListIzm(string delete)
        {
            string s = "";
            using (var reader = new StreamReader($"{AppDomain.CurrentDomain.BaseDirectory}/Raspisanie/allizm.txt"))
            {
                s = reader.ReadToEnd();
            }
            return Ok(s);
        }
    }
}
