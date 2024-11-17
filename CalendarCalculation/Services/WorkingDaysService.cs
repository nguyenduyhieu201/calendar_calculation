using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalendarCalculation.Services
{
    public static class WorkingDaysService
    {
        //tính số ngày làm việc trong tháng
        public static int WorkingDaysInMonth(int year, int month, DateTime startDate, DateTime endDate)
        {
            var listHoliday = new List<DateTime>
            {
                new DateTime(2024,1,1),
                new DateTime(2024,2,8),
                new DateTime(2024,2,9),
                new DateTime(2024,2,10),
                new DateTime(2024,2,11),
                new DateTime(2024,2,12),
                new DateTime(2024,2,13),
                new DateTime(2024,2,14),
                new DateTime(2024,4,18),
                new DateTime(2024,4,30),
                new DateTime(2024,5,1),
                new DateTime(2024,9,2),
                new DateTime(2024,9,3),
            };
            // Get the number of days in the month
            int numDays = DateTime.DaysInMonth(year, month);

            // Initialize a counter for working days
            int workingDays = 0;

            // Iterate through each day in the month
            for (int day = 1; day <= numDays; day++)
            {
                DateTime date = new DateTime(year, month, day);
                if (month == startDate.Month && day < startDate.Day)
                {
                    continue;
                }

                if (month == endDate.Month && day > endDate.Day)
                {
                    continue;
                }

                // Check if the day is a weekday (Monday to Friday)
                if (date.DayOfWeek != DayOfWeek.Saturday && date.DayOfWeek != DayOfWeek.Sunday && !listHoliday.Any(holiday => holiday.Equals(date)))
                {
                    workingDays++;
                }
            }

            return workingDays;
        }
        //tính số ngày làm việc trong một quãng thời gian, chia ra theo tháng
        public static Dictionary<Tuple<int, int>, Tuple<int, int>> WorkingDaysInDuration(DateTime startDate, DateTime endDate)
        {
            var workWatch = Stopwatch.StartNew();

            // Initialize a dictionary to store the count of working days for each month
            Dictionary<Tuple<int, int>, Tuple<int, int>> workingDaysPerMonth = new Dictionary<Tuple<int, int>, Tuple<int, int>>();

            // Iterate through each month within the duration
            while (startDate <= endDate)
            {
                int year = startDate.Year;
                int month = startDate.Month;
                // Calculate the number of working days in the current month
                int workingDays = WorkingDaysInMonth(year, month, startDate, endDate);
                int expectedWorkingHours = workingDays * 8;
                // Store the count in the dictionary
                workingDaysPerMonth.Add(new Tuple<int, int>(year, month), new Tuple<int, int>(workingDays, expectedWorkingHours));
                // Move to the next month
                startDate = new DateTime(year, month, 1).AddMonths(1);
            }

            return workingDaysPerMonth;
        }


    }
}
