using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EPPlusExam.Infrastructure
{
    public static class Utilities
    {
        public static bool IsNumeric(this string value)
        {
            bool isNum;
            double retNum;
            isNum = double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum;
        }
    }
}