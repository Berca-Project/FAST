using System;
using System.Globalization;
using System.IO;
using System.Web;

namespace Fast.Web.Utils
{
    public static class Helper
    {
        public static int GetCurrentShift()
        {
            // get shift based on current time
            if (Int32.Parse(DateTime.Now.Hour.ToString()) >= 6 && Int32.Parse(DateTime.Now.Hour.ToString()) < 14)
            {
                return 1;
            }
            else if (Int32.Parse(DateTime.Now.Hour.ToString()) >= 14 && Int32.Parse(DateTime.Now.Hour.ToString()) < 22)
            {
                return 2;
            }
            else
            {
                if (Int32.Parse(DateTime.Now.Hour.ToString()) >= 0 && Int32.Parse(DateTime.Now.Hour.ToString()) < 6)
                {
                    // 4 refers to shift 3 with previous date
                    return 4;
                }

                return 3;
            }
        }

        public static DateTime FirstDateOfWeek(string yearParam, string weekOfYearParam)
        {
            int year = int.Parse(yearParam);
            int weekOfYear = int.Parse(weekOfYearParam);

            DateTime jan1 = new DateTime(year, 1, 1);

            int daysOffset = (int)DayOfWeek.Monday - (int)jan1.DayOfWeek;

            DateTime firstMonday = jan1.AddDays(daysOffset);

            int firstWeek = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(jan1, CultureInfo.CurrentCulture.DateTimeFormat.CalendarWeekRule, DayOfWeek.Monday);

            if (firstWeek <= 1)
            {
                weekOfYear -= 1;
            }

            return firstMonday.AddDays(weekOfYear * 7);
        }

        public static DateTime LastDateOfWeek(string yearParam, string weekOfYearParam)
        {
            int year = int.Parse(yearParam);
            int weekOfYear = int.Parse(weekOfYearParam);

            DateTime jan1 = new DateTime(year, 1, 1);

            int daysOffset = (int)DayOfWeek.Monday - (int)jan1.DayOfWeek;

            DateTime firstMonday = jan1.AddDays(daysOffset);

            int firstWeek = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(jan1, CultureInfo.CurrentCulture.DateTimeFormat.CalendarWeekRule, DayOfWeek.Monday);

            if (firstWeek <= 1)
            {
                weekOfYear -= 1;
            }

            return firstMonday.AddDays(weekOfYear * 7).AddDays(6);
        }

        public static int GetWeekNumber(DateTime date)
        {
            CultureInfo ciCurr = CultureInfo.CurrentCulture;
            int weekNum = ciCurr.Calendar.GetWeekOfYear(date, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);
            return weekNum;
        }

        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public static string Number2String(int number)
        {
            if (number <= 26)
            {
                Char c = (Char)((97) + (number - 1));
                string result = c.ToString().ToUpper() + "{0}";
                return result;
            }
            else
            {
                if (number == 27)
                    return "AA{0}";
                else if (number == 28)
                    return "AB{0}";
                else if (number == 29)
                    return "AC{0}";
                else if (number == 30)
                    return "AD{0}";
                else if (number == 31)
                    return "AE{0}";
                else if (number == 32)
                    return "AF{0}";
                else if (number == 33)
                    return "AG{0}";
                else if (number == 34)
                    return "AH{0}";
                else if (number == 35)
                    return "AI{0}";
                else if (number == 36)
                    return "AJ{0}";
                else if (number == 37)
                    return "AK{0}";
                else if (number == 38)
                    return "AL{0}";
                else if (number == 39)
                    return "AM{0}";
                else if (number == 40)
                    return "AN{0}";
                else if (number == 41)
                    return "AO{0}";
                else if (number == 42)
                    return "AP{0}";
                else if (number == 43)
                    return "AQ{0}";
                else if (number == 44)
                    return "AR{0}";
                else if (number == 45)
                    return "AS{0}";
                else if (number == 46)
                    return "AT{0}";
                else if (number == 47)
                    return "AU{0}";
                else if (number == 48)
                    return "AV{0}";
                else if (number == 49)
                    return "AW{0}";
                else if (number == 50)
                    return "AX{0}";
                else if (number == 51)
                    return "AY{0}";
                else if (number == 52)
                    return "AZ{0}";
                else if (number == 53)
                    return "BA{0}";
                else if (number == 54)
                    return "BB{0}";
                else if (number == 55)
                    return "BC{0}";
                else if (number == 56)
                    return "BD{0}";
                else if (number == 57)
                    return "BE{0}";
                else if (number == 58)
                    return "BF{0}";
                else if (number == 59)
                    return "BG{0}";
                else if (number == 60)
                    return "BH{0}";
                else if (number == 61)
                    return "BI{0}";
                else if (number == 62)
                    return "BJ{0}";
                else if (number == 63)
                    return "BK{0}";
                else if (number == 64)
                    return "BL{0}";
                else if (number == 65)
                    return "BM{0}";
                else if (number == 66)
                    return "BN{0}";
                else if (number == 67)
                    return "BO{0}";
                else if (number == 68)
                    return "BP{0}";
                else if (number == 69)
                    return "BQ{0}";
                else if (number == 70)
                    return "BR{0}";
                else if (number == 71)
                    return "BS{0}";
                else if (number == 72)
                    return "BT{0}";
                else if (number == 73)
                    return "BU{0}";
                else if (number == 74)
                    return "BV{0}";
                else if (number == 75)
                    return "BW{0}";
                else if (number == 76)
                    return "BX{0}";
                else if (number == 77)
                    return "BY{0}";
                else if (number == 78)
                    return "BZ{0}";
                else if (number == 79)
                    return "CA{0}";
                else if (number == 80)
                    return "CB{0}";
                else if (number == 81)
                    return "CC{0}";
                else if (number == 82)
                    return "CD{0}";
                else if (number == 83)
                    return "CE{0}";
                else if (number == 84)
                    return "CF{0}";
                else if (number == 85)
                    return "CG{0}";
                else if (number == 86)
                    return "CH{0}";
                else if (number == 87)
                    return "CI{0}";
                else if (number == 88)
                    return "CJ{0}";
                else if (number == 89)
                    return "CK{0}";
                else if (number == 90)
                    return "CL{0}";
                else if (number == 91)
                    return "CM{0}";
                else if (number == 92)
                    return "CN{0}";
                else if (number == 93)
                    return "CO{0}";
                else if (number == 94)
                    return "CP{0}";
                else if (number == 95)
                    return "CQ{0}";
                else if (number == 96)
                    return "CR{0}";
                else if (number == 97)
                    return "CS{0}";
                else if (number == 98)
                    return "CT{0}";
                else if (number == 99)
                    return "CU{0}";
                else if (number == 100)
                    return "CV{0}";
                else if (number == 101)
                    return "CW{0}";
                else if (number == 102)
                    return "CX{0}";
                else if (number == 103)
                    return "CY{0}";
                else if (number == 104)
                    return "CZ{0}";
                else if (number == 105)
                    return "DA{0}";
                else if (number == 106)
                    return "DB{0}";
                else if (number == 107)
                    return "DC{0}";
                else if (number == 108)
                    return "DD{0}";
                else if (number == 109)
                    return "DE{0}";
                else if (number == 110)
                    return "DF{0}";
                else if (number == 111)
                    return "DG{0}";
                else if (number == 112)
                    return "DH{0}";
                else if (number == 113)
                    return "DI{0}";
                else if (number == 114)
                    return "DJ{0}";
                else if (number == 115)
                    return "DK{0}";
                else if (number == 116)
                    return "DL{0}";
                else if (number == 117)
                    return "DM{0}";
                else if (number == 118)
                    return "DN{0}";
                else if (number == 119)
                    return "DO{0}";
                else if (number == 120)
                    return "DP{0}";
                else if (number == 121)
                    return "DQ{0}";
                else if (number == 122)
                    return "DR{0}";
                else if (number == 123)
                    return "DS{0}";
                else if (number == 124)
                    return "DT{0}";
                else if (number == 125)
                    return "DU{0}";
                else if (number == 126)
                    return "DV{0}";
                else if (number == 127)
                    return "DW{0}";
                else if (number == 128)
                    return "DX{0}";
                else if (number == 129)
                    return "DY{0}";
                else
                    return "DZ{0}";
            }
        }

        public static string GetBrowser()
        {
            return HttpContext.Current.Request.Browser.Browser + " " + HttpContext.Current.Request.Browser.Version;
        }

        public static string GetVisitorIPAddress()
        {
            string IPAdd = string.Empty;
            IPAdd = HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];
            if (string.IsNullOrEmpty(IPAdd))
                IPAdd = HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"];
            return IPAdd;
        }

        private static string RemoveExtraHyphen(string text)
        {
            if (text.Contains("__"))
            {
                text = text.Replace("__", "_");
                return RemoveExtraHyphen(text);
            }
            return text;
        }

        public static void LogErrorMessage(Exception ex, string dir)
        {
            LogErrorMessage(ex.ToString(), dir);
        }

        public static void LogErrorMessage(string error, string dir, string empID = null)
        {
            if (empID == null)
            {
                using (StreamWriter streamWriter = File.AppendText(dir + "/log.txt"))
                {
                    Log(error, streamWriter);
                    streamWriter.Close();
                }
            }
            else
            {
                using (StreamWriter streamWriter = File.AppendText(dir + "/" + empID + ".txt"))
                {
                    Log(error, streamWriter);
                    streamWriter.Close();
                }
            }
        }

        public static void LogErrorMessageAddback(string error, string dir, string empID = null)
        {
            if (empID == null)
            {
                using (StreamWriter streamWriter = File.AppendText(dir + "/json-log.txt"))
                {
                    Log(error, streamWriter);
                    streamWriter.Close();
                }
            }
            else
            {
                using (StreamWriter streamWriter = File.AppendText(dir + "/json-" + empID + ".txt"))
                {
                    Log(error, streamWriter);
                    streamWriter.Close();
                }
            }
        }

        private static void Log(String logMessage, TextWriter w)
        {
            w.Write("\r\nLog Entry : ");
            w.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
            w.WriteLine("Message : {0}", logMessage);
            w.WriteLine("-------------------------------------------");
            w.Flush();
        }
    }
}