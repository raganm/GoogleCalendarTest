using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Microsoft.Office.Interop.Excel;

namespace CalendarGenerator.Utils
{
    public class CalendarUtils
    {
        private readonly List<string> _holidayTerms;
        private readonly List<string> _ascTerms;
        public readonly string HockeyCalendarId = "n4kf0tn6qag3t6jqvfct3blei8@group.calendar.google.com";
        public readonly string RaganCalendarId = "mcgill.ragan@gmail.com";
        public readonly string HayleyCalendarId = "hayley.mcgill@googlemail.com";
        public readonly string BinCalendarId = "pthclls6qgesvg2iojmgj508vk@group.calendar.google.com";
        public readonly string BirthdayCalendarId = "pfnm9gobovns3lgd7799g0j3ps@group.calendar.google.com";
        private readonly List<McGillDay> _days = new List<McGillDay>();

        public CalendarUtils(List<string> holidayTerms, List<string> ascTerms)
        {
            _holidayTerms = holidayTerms;
            _ascTerms = ascTerms;
        }

        public List<McGillDay> GenerateDays(DateTime startDate, DateTime endDate, CalendarService service)
        {
            var numberOfDays = DateTime.DaysInMonth(startDate.Year, startDate.Month);

            for (var i = 0; i < numberOfDays; i++)
            {
                _days.Add(new McGillDay(startDate.AddDays(i)));
            }

            var eventRequester = new EventRequester(service);

            var hockeyEvents = eventRequester.GetEvents(startDate, endDate, HockeyCalendarId);
            var raganEvents = eventRequester.GetEvents(startDate, endDate, RaganCalendarId);
            var hayleyEvents = eventRequester.GetEvents(startDate, endDate, HayleyCalendarId);
            var binEvents = eventRequester.GetEvents(startDate, endDate, BinCalendarId);
            var birthdayEvents = eventRequester.GetEvents(startDate, endDate, BirthdayCalendarId);

            SetHockeyDays(hockeyEvents);
            //SetEvents(raganEvents);
            //SetEvents(hayleyEvents);
            
            var events = new List<Event>();
            events.AddRange(raganEvents.Items);
            events.AddRange(hayleyEvents.Items);
            var wrappedEvents = events.Select(gEvent => new GoogleEventWrapper(gEvent)).ToList();

            var sortedEvents = wrappedEvents.OrderBy(x => x.Start).ToList(); // ToList optional
            SetEvents(sortedEvents);    
            
            SetBinDays(binEvents);
            SetBirthdays(birthdayEvents);

            return _days;
        }

        public CalendarService CreateService(string applicationName)
        {
            UserCredential credential;
            string[] scopes = { CalendarService.Scope.CalendarReadonly };

            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                var credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);

                credPath = Path.Combine(credPath, ".credentials");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
            }

            var service = new CalendarService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = applicationName,
            });

            return service;
        }

        public void SetBinDays(Events binEvents)
        {
            foreach (var binEvent in binEvents.Items)
            {
                var googleEvent = new GoogleEventWrapper(binEvent);

                foreach (var mcgillDay in _days)
                {
                    if (mcgillDay.Date.Date >= googleEvent.Start && mcgillDay.Date.Date < googleEvent.End && binEvent.Summary.ToLower().Contains("green"))
                    {
                        mcgillDay.IsGreenBinWeek = true;
                    }
                }
            }
        }

        public void SetHockeyDays(Events hockeyEvents)
        {
            foreach (var hockeyEvent in hockeyEvents.Items)
            {
                var googleEvent = new GoogleEventWrapper(hockeyEvent);

                var day = _days.First(x => x.Date.Date >= googleEvent.Start.Date && x.Date.Date <= googleEvent.End.Date);

                if (hockeyEvent.Summary.ToLower().Contains("oliver"))
                {
                    day.OliverHockey.Add(new HockeyEvent(googleEvent.Start, googleEvent.End, hockeyEvent.Summary));
                }else if (hockeyEvent.Summary.ToLower().Contains("bradley"))
                {
                    day.BradleyHockey.Add(new HockeyEvent(googleEvent.Start, googleEvent.End, hockeyEvent.Summary));
                }
                else
                {
                    day.OliverHockey.Add(new HockeyEvent(googleEvent.Start, googleEvent.End, hockeyEvent.Summary));
                    day.BradleyHockey.Add(new HockeyEvent(googleEvent.Start, googleEvent.End, hockeyEvent.Summary));
                }

            }
        }

        public void SetEvents(List<GoogleEventWrapper> events)
        {
            foreach (var googleEvent in events)
            {
                //var googleEvent = new GoogleEventWrapper(generalEvent);

                var days = _days.Where(x => x.Date.Date >= googleEvent.Start.Date && x.Date.Date <= googleEvent.End.Date).ToList();

                foreach (var day in days)
                {
                    var summary = googleEvent.GoogleEvent.Summary ?? "NO TITLE";
                    summary = summary.Replace(Environment.NewLine, " ");
                    summary = summary.Replace("\n", " ");
                    if (_holidayTerms.Any(word => summary.ToLower().Contains(word)))
                    {
                        if (summary.ToLower().Contains("term") || summary.ToLower().Contains("holiday"))
                        {
                            day.SchoolHoliday = true;
                        }

                        if (summary.ToLower().Contains("ragan"))
                        {
                            day.RaganHoliday = true;
                        }

                        if (summary.ToLower().Contains("hayley"))
                        {
                            day.HayleyHoliday = true;
                        }

                        if (summary.ToLower().Contains("bradley"))
                        {
                            day.BradleyHoliday = true;
                        }

                        if (summary.ToLower().Contains("oliver"))
                        {
                            day.OliverHoliday = true;
                        }
                    }
                    else if (_ascTerms.Any(word => summary.Contains(word)))
                    {
                        var where = summary.ToLower();

                        foreach (var ascTerm in _ascTerms)
                        {
                            where = where.Replace(ascTerm, "");
                        }

                        day.AscWhere = where.Trim();
                    }
                    else
                    {
                        string description;
                        if (googleEvent.Start.TimeOfDay.Hours == 0 && googleEvent.Start.TimeOfDay.Minutes == 0)
                        {
                            description = string.Format("{0}", summary);
                        }
                        else
                        {
                            description = string.Format("{0} {1}", googleEvent.Start.ToString("HH:mm"), summary);
                        }

                        if (string.IsNullOrWhiteSpace(day.Events))
                        {
                            day.Events = description;
                        }
                        else
                        {
                            day.Events += string.Format("      {0}", description);
                        }
                    }
                }

            }
        }

        public void SetBirthdays(Events birthdayEvents)
        {
            foreach (var birthdayEvent in birthdayEvents.Items)
            {
                var googleEvent = new GoogleEventWrapper(birthdayEvent);

                foreach (var mcgillDay in _days)
                {
                    if (mcgillDay.Date.Date >= googleEvent.Start && mcgillDay.Date.Date < googleEvent.End)
                    {
                        mcgillDay.Birthdays.Add(new Birthday(birthdayEvent.Summary));
                    }
                }
            }
        }

        public void GenerateCalendar(DateTime startDate, List<McGillDay> days, Color holidayColour, Color ascColour)
        {
            var xlApp = new Application { Visible = true };
            var spreadsheetLocation = Path.Combine(Directory.GetCurrentDirectory(), "Calendar TEMPLATE.xlsx");

            var workbook = xlApp.Workbooks.Open(spreadsheetLocation);

            var oSheet = (_Worksheet)workbook.ActiveSheet;

            var calendarBuilder = new CalendarBuilder(oSheet);

            calendarBuilder.SetTitle(startDate.ToString("MMMM", CultureInfo.InvariantCulture));

            calendarBuilder.Create(days, holidayColour,ascColour);

            xlApp.Visible = true;
        }
    }
}
