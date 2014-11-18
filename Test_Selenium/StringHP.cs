using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test_Selenium
{
    public class StringHP
    {
        //public const string ShareFolder = @"//LEXBUILD1/AutoScreenShot";

    }

    public class Message
    {
        public const string ExamNotFound = "the required exam is not found";
        public const string FormNotFound = "the required form is not found";
        public const string Success = "element is found successfully";
        public const string WindowNotFound = "the required window is not found";
        public const string NodeNotFound = "the tests node is not found";
        public const string Invalid = "INVALID";
    }

    public class ItemType
    {
        public const string MC = "ITSMultipleChoice";
        public const string HA = "ITSHotAreaBody";
        public const string DD = "ITSDragDropBody";
        public const string AS = "ITSActiveScreenBody";
        public const string BL = "ITSBuildListBody";
    }

    public enum Browser
    {
        IE,
        chrome,
        Firefox
    }
}
