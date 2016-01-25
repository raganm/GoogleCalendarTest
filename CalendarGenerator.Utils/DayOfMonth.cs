using System;
using System.Collections.Generic;

namespace CalendarGenerator.Utils
{
    public class DayOfMonth
    {
        public DateTime Date { get; set; }
        public List<HockeyEvent> OliverHockey { get; set; }
        public List<HockeyEvent> BradleyHockey { get; set; }
        public List<Birthday> Birthdays { get; set; }
        public string Events { get; set; }
        public bool IsGreenBinWeek { get; set; }
        public bool IsSchoolHoliday { get; set; }
        public bool IsHayleyOnHoliday { get; set; }
        public bool IsRaganOnHoliday { get; set; }
        public bool IsBradleyOnHoliday { get; set; }
        public bool IsOliverOnHoliday { get; set; }
        public string AscWhere { get; set; }

        public DayOfMonth(DateTime date)
        {
            Date = date;
            Birthdays = new List<Birthday>();
            OliverHockey = new List<HockeyEvent>();
            BradleyHockey = new List<HockeyEvent>();
        }

        public int Day
        {
            get { return Date.Day; }
        }

        public string DayName
        {
            get { return Date.ToString("dddd"); }
        }

        public bool IsWeekend
        {
            get
            {
                return Date.DayOfWeek == DayOfWeek.Saturday || Date.DayOfWeek == DayOfWeek.Sunday;
            }
        }

        public string HolidayDescription
        {
            get
            {
                var description = string.Empty;
                
                if (IsRaganOnHoliday)
                {
                    description += "R";
                }
                if (IsHayleyOnHoliday)
                {
                    description += "H";
                }
                if (IsBradleyOnHoliday)
                {
                    //description += "B";
                }
                if (IsOliverOnHoliday)
                {
                    //description += "O";
                }

                return description;
            }
        }

        public bool IsChildcareNeeded
        {
            get
            {
                return (Date.DayOfWeek == DayOfWeek.Thursday || Date.DayOfWeek == DayOfWeek.Friday) &&
                       (IsRaganOnHoliday == false && IsHayleyOnHoliday == false) && IsSchoolHoliday;
            }
        }
    }
}
