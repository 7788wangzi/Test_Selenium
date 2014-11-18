using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test_Selenium
{
    public class GlobalExam
    {
        public static int FormsCount { get; set; }

        public static string Name { get; set; }

        public static IList<string> Windows { get; set; }

        public static string TestName { get; set; }

        public static IList<string> Results { get; set; }
    }

    public class GlobalSets
    {
        public static Browser CurrentBrower { get; set; }

        public static string ResultSavePath { get; set; }
    }
}
