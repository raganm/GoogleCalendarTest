using System;
using Google.Apis.Calendar.v3.Data;

namespace CalendarGenerator.Utils
{
    public class GoogleEventWrapper
    {
        public Event GoogleEvent { get; set; }

        public GoogleEventWrapper(Event birthdayEvent)
        {
            GoogleEvent = birthdayEvent;
        }

        public DateTime Start
        {
            get
            {
                if (GoogleEvent.Start.DateTime != null && GoogleEvent.End.DateTime != null)
                {
                    return (DateTime)GoogleEvent.Start.DateTime;
                }
                else
                {
                    return Convert.ToDateTime(GoogleEvent.Start.Date);
                }
            }
        }

        public DateTime End
        {
            get
            {
                if (GoogleEvent.Start.DateTime != null && GoogleEvent.End.DateTime != null)
                {
                    return (DateTime)GoogleEvent.End.DateTime;
                }
                else
                {
                    return Convert.ToDateTime(GoogleEvent.End.Date).AddMinutes(-1);
                }
            }
        }
    }
}
