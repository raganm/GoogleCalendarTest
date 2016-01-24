using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalendarGenerator.Utils
{
    public class BinEventHelper
    {
        public readonly string BinCalendarId = "pthclls6qgesvg2iojmgj508vk@group.calendar.google.com";
        private readonly EventRequester _eventRequester;

        public BinEventHelper(EventRequester eventRequester)
        {
            _eventRequester = eventRequester;
        }

        public List<DayOfMonth> GetEvents( DateTime startDate, DateTime endDate)
        {
            var events = _eventRequester.GetEvents(startDate, endDate, BinCalendarId) ;
            var days = new List<DayOfMonth>();

            var xx =  events.Items.Select(gEvent => new GoogleEventWrapper(gEvent)).ToList();

            var numberOfDays = DateTime.DaysInMonth(startDate.Year, startDate.Month);

            for (var i = 0; i < numberOfDays; i++)
            {
                var mD = new DayOfMonth(startDate.AddDays(i));

                mD.IsGreenBinWeek = xx.Select(x => x.Start <= mD.Date && x.End >= mD.Date).FirstOrDefault();

                days.Add(mD);
            }

            foreach (var dayOfMonth in days)
            {
                Console.WriteLine("Date : {0} | IsGreeBinWeek : {1}", dayOfMonth.Date.ToShortDateString(),dayOfMonth.IsGreenBinWeek);
            }

            return days;
        } 
    }
}
