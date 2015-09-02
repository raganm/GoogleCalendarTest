using System;

namespace CalendarGenerator.Utils
{
    public class HockeyEvent
    {
        public bool IsTraining;
        public bool IsGame { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
        public string Summary { get; set; }
        public string Text { get; set; }

        private string UppercaseFirst(string s)
        {
            // Check for empty string.
            if (string.IsNullOrEmpty(s))
            {
                return string.Empty;
            }
            // Return char and concat substring.
            return char.ToUpper(s[0]) + s.Substring(1);
        }

        public string Opponent
        {
            get
            {
                var opponent = Summary;

                opponent = opponent.ToLower().Replace("bradley", "");
                opponent = opponent.ToLower().Replace("oliver", "");
                opponent = opponent.ToLower().Replace("vs", "");
                opponent = opponent.ToLower().Replace("home", "");
                opponent = opponent.ToLower().Replace("away", "");
                opponent = opponent.ToLower().Replace("friendly", "");
                opponent = UppercaseFirst(opponent.Trim());

                return Summary.ToLower().Contains("friendly") ? string.Format("{0} (F)", opponent) : opponent;
            }
        }

        public bool IsHome
        {
            get { return Summary.ToLower().Contains("home"); }
        }

        public bool IsAway
        {
            get { return Summary.ToLower().Contains("away"); }
        }

        public HockeyEvent(DateTime eventStartDate, DateTime eventEndDate, string summary)
        {
            StartDate = eventStartDate;
            EndDate = eventEndDate;
            Summary = summary;

            if (Summary.ToLower().Contains("vs"))
            {
                IsGame = true;
                Text = string.Format("{0} {1}", Opponent, StartDate.ToString("HH:mm"));
            }
            else
            {
                IsTraining = true;
                Text = string.Format("{0} - {1}", StartDate.ToString("HH:mm"), EndDate.ToString("HH:mm"));
            }
        }
    }
}