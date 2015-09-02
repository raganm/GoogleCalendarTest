using System;
using System.Collections.Generic;
using System.Drawing;
using CalendarGenerator.Utils;

namespace GoogleCalendarTest
{
    class Program
    {
        static string ApplicationName = "Google Calendar API Quickstart";
        private static readonly List<string> _holidayTerms = new List<string> { "off", "holiday", "term" };
        private static readonly List<string> _ascTerms = new List<string> { "boys -", "boys-" };

        static void Main(string[] args)
        {
            var startDate = new DateTime(Convert.ToInt32(args[1]), Convert.ToInt32(args[0]),01);
            var endDate = startDate.AddDays(DateTime.DaysInMonth(startDate.Year, startDate.Month)).AddDays(-1);

            var calendarUtils = new CalendarUtils(_holidayTerms,_ascTerms);
            var service = calendarUtils.CreateService(ApplicationName);
            var days = calendarUtils.GenerateDays(startDate, endDate, service);
            calendarUtils.GenerateCalendar(startDate, days,Color.LimeGreen, Color.BurlyWood);

        }

    }
}
