using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CalendarCalculation.Common
{
    public static class CommonVariable
    {
        public const string DefaultCalendarSheetName = "5.Resource Allocation";
        public const string DefaultOTSheetName = "Data.OTdetail";
        public const string DefaultTMSSheetName = "Data.TMS";
        public const string DefaultAbnormalSheetName = "AbnormalCase";

        public static string CalendarSheetName = "Calendar";
        public static string OTSheetName = "OT";
        public static string TMSSheetName = "TMS";
        public static string AbnormalSheetName = "AbnormalCase";
        public const int PartialDayLeave = 0;
        public const int FullDayLeave = 1;

        public static string UserName = "";
        public static string Password = "";
        public static string InputFolder = "";
        public static string OutputFilePath = "";

        public static bool IsTest = false;
    }
}
