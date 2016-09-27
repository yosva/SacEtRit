using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SacEtRit
{
    public class SacReport : ISacReport
    {
        public void Create(double heures, string developpeur, string client, out int totalJouors, int? mois = null, string outPath = null)
        {
            MemoryStream ms = new MemoryStream(Resource1.Suivi_Activite_Mensuel);

            if (!mois.HasValue)
                mois = DateTime.Now.Month;

            DateTime dt = new DateTime(DateTime.Now.Year, mois.Value, 1);

            CultureInfo ci = new CultureInfo("fr-FR");

            string monthName = dt.ToString("MMMM", ci);

            if (String.IsNullOrEmpty(outPath))
                outPath = $"{AppDomain.CurrentDomain.BaseDirectory}Suivi_Activite_Mensuel_{monthName}.xlsx";
            else
                outPath = $"{outPath}\\Suivi_Activite_Mensuel_{monthName}.xlsx";

            double daily = heures / 5.0;

            int h = (int)daily;
            int m = (int)Math.Round((daily - (double)h)*60.0);
            TimeSpan ts = new TimeSpan(h, m, 0);
            totalJouors = 0;

            using (ExcelPackage package = new ExcelPackage(ms))
            {
                bool ready = false;
                ExcelWorksheet worksheet = null;
                do
                {
                    try
                    {
                        worksheet = package.Workbook.Worksheets["Mensuel"];//.First(); //peut être un bug dans EPPlus, la première fois qu'on essaye de lire le sheet une exception se lève avec message "an item with the same key has already been added"
                        ready = true;
                        ///TODO faire en sorte qu'on arrète la boucle si depuis une limiete de temps on ne peut toujours pas lire le sheet...jusqu'à présent ça n'as pas arrivé
                    }
                    catch (Exception)
                    {

                        //throw;
                        continue;
                    }
                }
                while (!ready);

                int firstWeekOfMonth = DateTimeFormatInfo.CurrentInfo.Calendar.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

                worksheet.Cells["D3"].Value = "NOM: " + developpeur;
                worksheet.Cells["E4"].Value = monthName;

                do
                {
                    if (dt.DayOfWeek != DayOfWeek.Saturday && dt.DayOfWeek != DayOfWeek.Sunday)
                    {
                        int week = DateTimeFormatInfo.CurrentInfo.Calendar.GetWeekOfYear(dt, CalendarWeekRule.FirstDay, DayOfWeek.Monday);

                        int weekOfMonth = week - firstWeekOfMonth + 1;

                        int I = 6 * weekOfMonth;

                        string weekCell = $"A{I}";
                        if(worksheet.Cells[weekCell].Value == null || string.IsNullOrEmpty(worksheet.Cells[weekCell].Value.ToString()))
                        {
                            worksheet.Cells[weekCell].Value = $"S{week}";
                        }

                        worksheet.Cells[I, (int)dt.DayOfWeek+1].Value = $"{dt.ToString("dddd", ci).ToUpper()} {dt.Day}";

                        worksheet.Cells[I + 2, (int)dt.DayOfWeek + 1].Value = $"MI ({client})";

                        worksheet.Cells[I + 3, (int)dt.DayOfWeek + 1].Style.Numberformat.Format = "[H]\"h\"MM";
                        worksheet.Cells[I + 3, (int)dt.DayOfWeek + 1].Value = ts;

                        worksheet.Cells[$"H{I + 2}"].Value = $"MI: {week}";
                        worksheet.Cells[$"H{I + 3}"].Style.Numberformat.Format = "[H]\"h\"MM";

                        ++totalJouors;
                    }

                    dt = dt.AddDays(1);
                } while (dt.Month == mois);

                worksheet.Cells["H36"].Style.Numberformat.Format = "[H]\"h\"MM";

                int M = (int)Math.Round((heures - Math.Truncate(heures)) * 60);
                string str = M==0 ? "" : M.ToString() + " minutes";
                worksheet.Cells["A37"].Value = $"Conformément aux contrats de travail, la durée hebdomadaire du travail est fixée à {(int)heures} heures {str} (soit {ts.Hours}h{ts.Minutes} par jour).";

                worksheet.Calculate();

                package.SaveAs(new FileInfo(outPath));
            }
        }
    }
}
