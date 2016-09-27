using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SacEtRit.Helpers
{
    public static class HolidaysHelper
    {
        /// <summary>
        ///     Gets the easter Date
        /// </summary>
        /// <param name="year">The year of the easter date</param>
        /// <returns></returns>
        private static DateTime EasterDate(int year)
        {
            var y = year;
            var a = y % 19;
            var b = y / 100;
            var c = y % 100;
            var d = b / 4;
            var e = b % 4;
            var f = (b + 8) / 25;
            var g = (b - f + 1) / 3;
            var h = (19 * a + b - d - g + 15) % 30;
            var i = c / 4;
            var k = c % 4;
            var l = (32 + 2 * e + 2 * i - h - k) % 7;
            var m = (a + 11 * h + 22 * l) / 451;
            var month = (h + l - 7 * m + 114) / 31;
            var day = ((h + l - 7 * m + 114) % 31) + 1;
            return new DateTime(year, month, day);
        }

        /// <summary>
        ///     Gets all french public holidays.
        /// </summary>
        /// <param name="year">The year.</param>
        /// <returns></returns>
        public static IEnumerable<HolidayViewModel> GetAll(int year)
        {
            var result = new List<HolidayViewModel>
            {
                //01/01
                new HolidayViewModel("Nouvel an", new DateTime(year, 1, 1)),
                //01/05
                new HolidayViewModel("Fête du travail", new DateTime(year, 5, 1)),
                //08/05
                new HolidayViewModel("Fête de la victoire", new DateTime(year, 5, 8)),
                //14/07
                new HolidayViewModel("Fête nationale", new DateTime(year, 7, 14)),
                //15/08
                new HolidayViewModel("Assomption", new DateTime(year, 8, 15)),
                //01/11
                new HolidayViewModel("Toussaint", new DateTime(year, 11, 1)),
                //11/11
                new HolidayViewModel("Armistice", new DateTime(year, 11, 11)),
                //25/12
                new HolidayViewModel("Noël", new DateTime(year, 12, 25)),
                //Easter Monday
                new HolidayViewModel("Lundi de Pâques", EasterDate(year).AddDays(1)),
                //Ascension
                new HolidayViewModel("Jeudi de l'Ascension", EasterDate(year).AddDays(39)),
                //Pentecôte
                new HolidayViewModel("Lundi de Pentecôte", EasterDate(year).AddDays(50))
            };

            return result;
        }

        /// <summary>
        /// Gets holidays between two dates.
        /// </summary>
        /// <param name="start">The start.</param>
        /// <param name="end">The end.</param>
        /// <returns></returns>
        public static IEnumerable<HolidayViewModel> GetRangeDate(DateTime start, DateTime end)
        {
            var holidays = new List<HolidayViewModel>();

            for (var i = start.Year; i <= end.Year; i++)
            {
                holidays.AddRange(GetAll(i));
            }

            return holidays.Where(w => w.Date >= start && w.Date <= end);
        }

        public struct HolidayViewModel
        {
            public HolidayViewModel(string name, DateTime date) : this()
            {
                Name = name;
                Date = date;
            }

            public string Name { get; set; }
            public DateTime Date { get; set; }
        }
    }
}
