using Aspose.Cells;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using Spire.Xls;
using System.Threading;
using System.Threading.Tasks;

namespace wbExample.Models
{
    public class Differents
    {
        public static WebClient webClient = new WebClient();
        public static List<string> dataforcb = new List<string>();
        public static XLWorkbook _workbook;
        public static IXLWorksheet workSheet;
        public static int upDay = 0;
        public static int downDay = 0;
        public static DateTime DupDay;
        public static DateTime DdownDay;
        public static string downMonth;
        public static string upMonth;
        public static void DateOut(DateTime dt)
        {
            switch (dt.DayOfWeek)
            {
                case DayOfWeek.Sunday:
                    CleanCache();
                    downDay = dt.AddDays(1).Day;
                    upDay = dt.AddDays(6).Day;
                    DupDay = dt.AddDays(6);
                    DdownDay = dt.AddDays(1);
                    break;
                case DayOfWeek.Monday:
                    downDay = dt.Day;
                    upDay = dt.AddDays(5).Day;
                    DupDay = dt.AddDays(5);
                    DdownDay = dt;
                    break;
                case DayOfWeek.Tuesday:
                    downDay = dt.AddDays(-1).Day;
                    upDay = dt.AddDays(4).Day;
                    DupDay = dt.AddDays(4);
                    DdownDay = dt.AddDays(-1);
                    break;
                case DayOfWeek.Wednesday:
                    downDay = dt.AddDays(-2).Day;
                    upDay = dt.AddDays(3).Day;
                    DupDay = dt.AddDays(3);
                    DdownDay = dt.AddDays(-2);
                    break;
                case DayOfWeek.Thursday:
                    downDay = dt.AddDays(-3).Day;
                    upDay = dt.AddDays(2).Day;
                    DupDay = dt.AddDays(2);
                    DdownDay = dt.AddDays(-3);
                    break;
                case DayOfWeek.Friday:
                    downDay = dt.AddDays(-4).Day;
                    upDay = dt.AddDays(1).Day;
                    DupDay = dt.AddDays(1);
                    DdownDay = dt.AddDays(-4);
                    break;
                case DayOfWeek.Saturday:
                    downDay = dt.AddDays(-5).Day;
                    upDay = dt.Day;
                    DupDay = dt;
                    DdownDay = dt.AddDays(-5);
                    break;
            }
            downMonth = Mouth(DdownDay);
            upMonth = Mouth(DupDay);
        }

        public static void CleanCache()
        {
            for (int i = 0; i < 6; i++)
            {
                try
                {
                    System.IO.File.Delete($"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/{DdownDay.AddDays(i - 7).ToShortDateString()}.xls");
                    System.IO.File.Delete($"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/{DdownDay.AddDays(i - 7).ToShortDateString()}.xlsx");
                }
                catch
                {
                    continue;
                }
            }
            try
            {
                System.IO.File.Delete($"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/_{DdownDay.AddDays(-7).ToShortDateString()}_{DupDay.AddDays(-7).ToShortDateString()}.xls");
            }
            catch
            {

            }
        }
        public static string Mouth(DateTime date)
        {
            var lines = System.IO.File.ReadAllText($"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/Month.txt");
            string[] massiv = lines.Split('\n');
            switch (date.Month)
            {
                case 9:
                    return massiv[8];
                case 10:
                    return massiv[9];
                case 11:
                    return massiv[10];
                case 12:
                    return massiv[11];
                case 1:
                    return massiv[0];
                case 2:
                    return massiv[1];
                case 3:
                    return massiv[2];
                case 4:
                    return massiv[3];
                case 5:
                    return massiv[4];
                case 6:
                    return massiv[5];
                case 7:
                    return massiv[6];
                case 8:
                    return massiv[7];
                default:
                    break;
            }
            return null;
        }
        public static void DownloadFeatures(DateTime time)
        {
            string data = "";
            using (var stream = new StreamReader($"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/allizm.txt"))
            {
                data = stream.ReadToEnd();
                stream.Close();
            }





            DateTime dateIZM = time;
            int dayWeek = (int)time.DayOfWeek;
            WebClient web = new WebClient();
            Aspose.Cells.Workbook workbook;
            CultureInfo culture = new CultureInfo("ru-RU");

            for (int i = 1; i <= dayWeek + 1; i++)
            {
                //if(File.Exists($"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/{dateIZM.AddDays(i - dayWeek).ToString("d", culture)}.xlsx"))
                if (data.Contains($"{dateIZM.AddDays(i - dayWeek).ToString("d", culture)}.xlsx"))
                {
                    continue;
                }
                else
                {
                    try
                    {
                        web.DownloadFile(@$"https://oksei.ru/files/{dateIZM.AddDays(i - dayWeek).ToString("d", culture)}.xls", @$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/{ dateIZM.AddDays(i - dayWeek).ToString("d", culture)}.xls");
                        Spire.Xls.Workbook workbook2 = new Spire.Xls.Workbook();
                        workbook2.LoadFromFile(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/{ dateIZM.AddDays(i - dayWeek).ToString("d", culture)}.xls");
                        workbook2.SaveToFile(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/{ dateIZM.AddDays(i - dayWeek).ToString("d", culture)}.xlsx", ExcelVersion.Version2013);
                        XLWorkbook xL = new XLWorkbook(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/{ dateIZM.AddDays(i - dayWeek).ToString("d", culture)}.xlsx");
                        Controllers.LastDanceController.RaspisanieIzm(xL, i);
                        _workbook.Save();
                        data += $"{dateIZM.AddDays(i - dayWeek).ToString("d", culture)}.xlsx\n";
                        using(StreamWriter streamWriter = new StreamWriter($"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/allizm.txt", false, System.Text.Encoding.Default))
                        {
                            streamWriter.Write(data);
                            streamWriter.Close();
                        }
                    }
                    catch
                    {
                        try
                        {
                            web.DownloadFile(@$"https://oksei.ru/files/{dateIZM.AddDays(i - dayWeek).Day}.{dateIZM.AddDays(i - dayWeek).Month}.{dateIZM.AddDays(i - dayWeek).Year}.xls", @$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/{ dateIZM.AddDays(i - dayWeek).ToString("d", culture)}.xls");
                            Spire.Xls.Workbook workbook2 = new Spire.Xls.Workbook();
                            workbook2.LoadFromFile(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/{ dateIZM.AddDays(i - dayWeek).ToString("d", culture)}.xls");
                            workbook2.SaveToFile(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/{ dateIZM.AddDays(i - dayWeek).ToString("d", culture)}.xlsx", ExcelVersion.Version2013);
                            XLWorkbook xL = new XLWorkbook(@$"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/{ dateIZM.AddDays(i - dayWeek).ToString("d", culture)}.xlsx");
                            Controllers.LastDanceController.RaspisanieIzm(xL, i);
                            _workbook.Save();
                            data += $"{dateIZM.AddDays(i - dayWeek).ToString("d", culture)}.xlsx\n";
                            using (StreamWriter streamWriter = new StreamWriter($"{AppDomain.CurrentDomain.BaseDirectory}Raspisanie/allizm.txt", false, System.Text.Encoding.Default))
                            {
                                streamWriter.Write(data);
                                streamWriter.Close();
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }                
            }
        }
        public static List<DayWeekClass> EnumerableMetrics(int row, int column)
        {
            int count = 1;
            List<DayWeekClass> dayWeeks = new List<DayWeekClass>();
            XLWorkbook workbook = _workbook;
            var worksheet = workbook.Worksheets.Worksheet(1);
            if (row > 24 && workSheet.Cell(27, 2).Value.ToString() != "4")
            {
                row++;
            }
            bool g = true;
            for (int i = row; i < row + 6; i++)
            {

                if (row == 24 && i == 27 && g)
                {
                    
                    if (worksheet.Cell(i, column).Value.ToString().Contains("ЧКР"))
                    {
                        row++;
                        dayWeeks.Add(new DayWeekClass { Number = null, Day = worksheet.Cell(i, column).GetValue<string>() });
                        continue;
                    }
                    else
                    {
                        dayWeeks.Add(new DayWeekClass { Number = null, Day = "ЧКР" });
                        i--;
                        g = false;
                        continue;
                    }
                }
                var metric = new DayWeekClass
                {
                    Number = count,
                    Day = worksheet.Cell(i, column).GetValue<string>()
                };
                count++;
                dayWeeks.Add(metric);

            }
            return dayWeeks;
        }
        public static int IndexGroup(string group)
        {
            var workSheet = _workbook.Worksheets.First();
            for (int i = 1; i < workSheet.ColumnCount(); i++)
            {
                if (workSheet.Cell(5, i).GetValue<string>() == group)
                {
                    return i;
                }
                else
                    continue;
            }
            return 0;
        }

        public static List<DayWeekClass> raspisaniekab(int row, string kabinet)
        {
            bool exit = false;
            int number = 1;
            List<DayWeekClass> kabinets = new List<DayWeekClass>();
            if (row > 24 && workSheet.Cell(27, 2).Value.ToString() != "4")
            {
                row++;
            }
            for (int i = row; i < row + 6; i++)
            {

                if (row == 24 && i == 27 )
                {
                    if (workSheet.Cell(27, 2).Value.ToString() != "4")
                    {
                        row++;
                        kabinets.Add(new DayWeekClass { Day = "ЧКР" });
                        continue;
                    }
                    else
                        kabinets.Add(new DayWeekClass { Day = "ЧКР" });
                }
                

                for (int j = 3; j <= workSheet.ColumnsUsed().Count(); j++)
                {
                    string result = workSheet.Cell(i, j).GetValue<string>();
                    if (result.Contains(kabinet))
                    {
                        kabinets.Add(new DayWeekClass { Number = number, Day = result + $"\n{workSheet.Cell(5, j).GetValue<string>()}" });
                        exit = false;
                        break;
                    }
                    else
                        exit = true;
                }
                if (exit)
                    kabinets.Add(new DayWeekClass { Number = number, Day = "-" });
                number++;
            }
            return kabinets;
        }
        public static List<DayWeekClass> raspisanieteach(int row, string teach)
        {

            bool exit = false;
            int number = 1;
            List<DayWeekClass> kabinets = new List<DayWeekClass>();
            if (row > 24 && workSheet.Cell(27, 2).Value.ToString() != "4")
            {
                row++;
            }
            for (int i = row; i < row + 6; i++)
            {

                if (row == 24 && i == 27)
                {
                    if (workSheet.Cell(27, 2).Value.ToString() != "4")
                    {
                        row++;
                        kabinets.Add(new DayWeekClass { Day = "ЧКР" });
                        continue;
                    }
                    else
                        kabinets.Add(new DayWeekClass { Day = "ЧКР" });
                }
                for (int j = 3; j <= workSheet.ColumnsUsed().Count(); j++)
                {
                    string result = workSheet.Cell(i, j).GetValue<string>();
                    if (result.Contains(teach))
                    {
                        kabinets.Add(new DayWeekClass { Number = number, Day = result + $"\n{workSheet.Cell(5, j).GetValue<string>()}" });
                        exit = false;
                        break;
                    }
                    else
                        exit = true;
                }
                if (exit)
                    kabinets.Add(new DayWeekClass { Number = number, Day = "-" });
                number++;
            }
            return kabinets;
        }
    }
}
