using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using Google.Apis.Calendar.v3.Data;
using Microsoft.Office.Interop.Excel;

namespace CalendarGenerator.Utils
{
    public class CalendarBuilder
    {
        private readonly _Worksheet _sheet;

        public CalendarBuilder(_Worksheet sheet)
        {
            _sheet = sheet;
        }

        public void SetTitle(string title)
        {
            _sheet.Cells[1, 1] = title;
        }

        public void Create(List<McGillDay> days, Color holidayColour, Color ascColor)
        {
            var startRow = 3;
            foreach (var mcgillDay in days)
            {
                _sheet.Cells[startRow, 1] = mcgillDay.Day;
                _sheet.Cells[startRow, 2] = mcgillDay.DayName.ToUpper();
                if (mcgillDay.IsGreenBinWeek)
                {
                    _sheet.Range["B" + startRow].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.MediumSeaGreen); ;
                }

                if (mcgillDay.IsAfterSchoolClub)
                {
                    _sheet.Range["C" + startRow].Interior.Color = System.Drawing.ColorTranslator.ToOle(ascColor);
                }

                _sheet.Cells[startRow, 3] = mcgillDay.AscWhere;

                foreach (var birthday in mcgillDay.Birthdays)
                {
                    _sheet.Cells[startRow, 4] = birthday.Name;
                }

                if (mcgillDay.OliverHockey != null)
                {
                    var summary = string.Empty;

                    foreach (var hockeyEvent in mcgillDay.OliverHockey)
                    {
                        if (summary == string.Empty)
                        {
                            summary = hockeyEvent.Text.Trim();
                        }
                        else
                        {
                            summary += "\n" + hockeyEvent.Text.Trim();
                        }

                        if (hockeyEvent.IsGame && hockeyEvent.IsHome)
                        {
                            _sheet.Range["E" + startRow].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.MediumSeaGreen); ;
                        }

                        if (hockeyEvent.IsGame && hockeyEvent.IsAway)
                        {
                            _sheet.Range["E" + startRow].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Crimson); ;
                        }
                    }

                    if (mcgillDay.OliverHockey.Count > 1)
                    {
                        _sheet.Range["E" + startRow].Font.Size = 8;
                    }

                    _sheet.Cells[startRow, 5].Value = summary;

                }

                if (mcgillDay.BradleyHockey != null)
                {
                    var summary = string.Empty;

                    foreach (var hockeyEvent in mcgillDay.BradleyHockey)
                    {
                        if (summary == string.Empty)
                        {
                            summary = hockeyEvent.Text;
                        }
                        else
                        {
                            summary += "\n" + hockeyEvent.Text;
                        }

                        if (hockeyEvent.IsGame && hockeyEvent.IsHome)
                        {
                            _sheet.Range["F" + startRow].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.MediumSeaGreen); ;
                        }

                        if (hockeyEvent.IsGame && hockeyEvent.IsAway)
                        {
                            _sheet.Range["F" + startRow].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Crimson); ;
                        }
                    }

                    if (mcgillDay.BradleyHockey.Count > 1)
                    {
                        _sheet.Range["F" + startRow].Font.Size = 8;
                    }

                    _sheet.Cells[startRow, 6].Value = summary;
                }

                if (mcgillDay.SchoolHoliday)
                {
                    _sheet.Range["G" + startRow].Interior.Color = System.Drawing.ColorTranslator.ToOle(holidayColour);
                    _sheet.Range["G" + startRow].Cells.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                }
                _sheet.Cells[startRow, 7] = mcgillDay.HolidayDescription;

                _sheet.Cells[startRow, 8] = mcgillDay.Events;
                startRow++;
            }
            _sheet.Range["A1"].Select();
        }
    }
}