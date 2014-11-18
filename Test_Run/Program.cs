using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Test_Selenium;
using System.IO;
using System.Xml;

namespace Test_Run
{
    class Program
    {
        /// <summary>
        /// Sample:
        /// Test_Run.exe "MB2-866-ENU-100" "MB2-720-ENU-100" "MB7-702-ENU-100"
        /// </summary>
        /// <param name="args">"MB2-866-ENU-100" "MB2-720-ENU-100" "MB7-702-ENU-100"</param>
        static void Main(string[] args)
        {
            string browserID = "1";
            string savePath = "";
            ReadConfig(out browserID, out savePath);
            if (savePath == Message.Invalid)
                return;
            List<string> examsToRun = new List<string>();
            foreach (var item in args)
            {
                examsToRun.Add(item);
            }
            //examsToRun.Add("AS-CS Batch 01 Content Review");
            examsToRun.Add("642-ENU-665");
            //examsToRun.AddRange(new string[] { "MB7-702-ENU-100", "MB7-701-ENU-100", "MB6-889-ENU-100" });
            //examsToRun.AddRange(new string[] { "MB6-884-ENU-100", "MB6-885-ENU-100", "MB6-886-ENU-101" });
//            examsToRun.AddRange(new string[] {
//            "MB2-866-ENU-100","MB2-720-ENU-100","MB2-703-ENU-100","MB2-702-ENU-100","MB2-701-ENU-100","MB2-700-ENU-100"});
//            examsToRun.AddRange(new string[] {"74-344-ENU-100","518-ENU-400","460-ENU-200","461-ENU-152","158-ENU-100","74-338-ENU-100","74-697-ENU-100","74-353-ENU-100",
//"74-674-ENU-200","74-335-ENU-101","62-193-ENU-200","177-ENU-100","178-ENU-200","346-ENU-101","347-ENU-100"});
            Class1 cls1 = new Class1();
            foreach (var item in examsToRun)
            {                
                Console.WriteLine(string.Format("Running exam {0}",item));
                try
                {
                    switch (browserID)
                    {
                        case "1": cls1.Run(item, Browser.IE, savePath); break;
                        case "2": cls1.Run(item, Browser.chrome, savePath); break;
                        case "3": cls1.Run(item, Browser.Firefox, savePath); break;
                        default: cls1.Run(item, Browser.IE, savePath); break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
                //cls1.Run(item,Browser.IE);
            }
            
        }

        static void ReadConfig(out string browserID, out string savePath)
        {
            string configFile = @"Test_Run_Config.xml";
            if (!File.Exists(configFile))
            {
                Console.WriteLine("Test_Run_Config.xml not found, make sure the config file exist please.");
                browserID = "1";
                savePath = Message.Invalid;
                return;
            }
            
            XmlDocument _xmlDoc = new XmlDocument();
            _xmlDoc.Load(configFile);

            XmlElement root = _xmlDoc.DocumentElement;
            browserID = root.SelectSingleNode(@"//Browser").InnerText;
            savePath = root.SelectSingleNode(@"//ScreenShot").InnerText;
        }
    }
}
