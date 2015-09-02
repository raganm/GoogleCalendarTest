using System;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;

namespace CalendarGenerator.Utils
{
    public class EventRequester
    {
        private readonly CalendarService _service;

        public EventRequester(CalendarService service)
        {
            _service = service;
        }

        public Events GetEvents(DateTime startDate, DateTime endDate, string calendarId)
        {
            // Define parameters of request.
            var request = _service.Events.List(calendarId);
            request.TimeMin = startDate;
            request.TimeMax = endDate;
            request.ShowDeleted = false;
            request.SingleEvents = true;
            //request.MaxResults = 100;
            request.OrderBy = EventsResource.ListRequest.OrderByEnum.StartTime;

            var events = request.Execute();
            
            return events;
        }
    }
}