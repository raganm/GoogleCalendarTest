using System;
using System.Collections.Generic;

namespace CalendarGenerator.Utils
{
    public class McGillDay
    {
        public DateTime Date { get; set; }
        public List<HockeyEvent> OliverHockey { get; set; }
        public List<HockeyEvent> BradleyHockey { get; set; }
        public List<Birthday> Birthdays { get; set; }
        public string Events { get; set; }
        public bool IsGreenBinWeek { get; set; }
        public bool SchoolHoliday { get; set; }
        public bool HayleyHoliday { get; set; }
        public bool RaganHoliday { get; set; }
        public bool BradleyHoliday { get; set; }
        public bool OliverHoliday { get; set; }
        public string AscWhere { get; set; }

        public McGillDay(DateTime date)
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

        public bool IsAfterSchoolClub
        {
            get
            {
                return false;
                //return !SchoolHoliday && (Date.DayOfWeek == DayOfWeek.Thursday) && (HayleyHoliday == false ||RaganHoliday == false);
                //return !SchoolHoliday && (Date.DayOfWeek == DayOfWeek.Thursday || Date.DayOfWeek == DayOfWeek.Friday) && (HayleyHoliday == false ||RaganHoliday == false);
                //return (Date.DayOfWeek == DayOfWeek.Thursday || Date.DayOfWeek == DayOfWeek.Friday) && (HayleyHoliday == false ||RaganHoliday == false);
            }
        }

        public string HolidayDescription
        {
            get
            {
                var description = string.Empty;
                
                if (RaganHoliday)
                {
                    description += "R";
                }
                if (HayleyHoliday)
                {
                    description += "H";
                }
                if (BradleyHoliday)
                {
                    //description += "B";
                }
                if (OliverHoliday)
                {
                    //description += "O";
                }

                return description;
            }
        }

    }
}
