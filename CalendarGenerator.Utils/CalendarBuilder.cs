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
                    _sheet.Cells[startRow, 5] = mcgillDay.OliverHockey.Text;

                    if (mcgillDay.OliverHockey.IsGame && mcgillDay.OliverHockey.IsHome)
                    {
                        _sheet.Range["E" + startRow].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.MediumSeaGreen); ;
                    }

                    if (mcgillDay.OliverHockey.IsGame && mcgillDay.OliverHockey.IsAway)
                    {
                        _sheet.Range["E" + startRow].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Crimson); ;
                    }
                }

                if (mcgillDay.BradleyHockey != null)
                {
                    _sheet.Cells[startRow, 6] = mcgillDay.BradleyHockey.Text;

                    if (mcgillDay.BradleyHockey.IsGame && mcgillDay.BradleyHockey.IsHome)
                    {
                        _sheet.Range["F" + startRow].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.MediumSeaGreen); ;
                    }

                    if (mcgillDay.BradleyHockey.IsGame && mcgillDay.BradleyHockey.IsAway)
                    {
                        _sheet.Range["F" + startRow].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Crimson); ;
                    }
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