using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;

namespace ExcelDnaTest
{

    public class UdfTest
    {
        [ExcelFunction(Name = "aa")]
        public static object aa()
        {
            var xlCall = XlCall.Excel(XlCall.xlGetInst);
            return xlCall;
        }

        [ExcelFunction(Name = "bb")]
        public static object bb()
        {
            var xlCall = XlCall.Excel(XlCall.xlGetHwnd);
            return xlCall;
        }

        [ExcelFunction(Name = "cc")]
        public static object cc()
        {
            var xlCall = XlCall.Excel(XlCall.xlModeEdit);
            return xlCall;
        }
    }
}
