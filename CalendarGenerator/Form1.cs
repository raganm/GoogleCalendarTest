using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using CalendarGenerator.Utils;

namespace CalendarGenerator
{
    public partial class Form1 : Form
    {
        private readonly string _applicationName;
        private readonly List<string> _holidayTerms = new List<string> { "off", "holiday", "term" };
        private readonly List<string> _ascTerms = new List<string> { "boys -", "boys-" };
        private Color _holidayColour;
        private Color _ascColour;


        public Form1()
        {
            InitializeComponent();

            txtMonth.Text = DateTime.Now.AddMonths(1).Month.ToString();
            txtYear.Text = DateTime.Now.Year.ToString();
            txtHolidayTerms.Text = string.Join(",",_holidayTerms);
            txtAscTerms.Text = string.Join(",",_ascTerms);

            SetHolidayColour(Color.LimeGreen);
            SetAscColour(Color.LightBlue);

            _applicationName = "Hayleys Calendar Builder";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            char[] charSeparators = new char[] { ',' };
            button1.Enabled = false;

            var startDate = new DateTime(Convert.ToInt32(txtYear.Text), Convert.ToInt32(txtMonth.Text), 01);
            var endDate = startDate.AddDays(DateTime.DaysInMonth(startDate.Year, startDate.Month)).AddDays(-1);

            var holidayTerms = txtHolidayTerms.Text.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries).ToList();
            var ascTerms = txtAscTerms.Text.Split(charSeparators, StringSplitOptions.RemoveEmptyEntries).ToList();

            var calendarUtils = new CalendarUtils(holidayTerms, ascTerms);
            var service = calendarUtils.CreateService(_applicationName);
            var days = calendarUtils.GenerateDays(startDate, endDate, service);
            calendarUtils.GenerateCalendar(startDate, days, _holidayColour,_ascColour);
            button1.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var result = colorDialog1.ShowDialog();

            if (result == DialogResult.OK)
            {
                SetHolidayColour(colorDialog1.Color);
            }
        }

        private void SetHolidayColour(Color color)
        {
            btnHolidayColour.BackColor = color;
            _holidayColour = color;
        }

        private void SetAscColour(Color color)
        {
            btnAscColour.BackColor = color;
            _ascColour = color;
        }

        private void btnAscColour_Click(object sender, EventArgs e)
        {
            var result = colorDialog1.ShowDialog();
            
            if (result == DialogResult.OK)
            {
                SetAscColour(colorDialog1.Color);
            }
        }
    }
}
