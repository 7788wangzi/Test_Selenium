using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;
using System.Text.RegularExpressions;
using System.IO;

namespace Test_Selenium
{
    public class KeysHP
    {
        public static ArrayList GetKeysForBL(string input)
        {
            ArrayList resultList = new ArrayList();
            // remove only
            string _only = "only";
            int _indexOfOnly = input.ToLower().IndexOf(_only);

            string validString = input;
            if (_indexOfOnly != -1)
            {
                validString = input.ToLower().Substring(0, _indexOfOnly);
            }

            // split by and            
            string[] values = validString.ToLower().Split(new string[] { "and" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var value in values)
            {
                if (string.IsNullOrWhiteSpace(value))
                    continue;
                string[] keys = value.Split(new string[] { "before" }, StringSplitOptions.RemoveEmptyEntries);
                if (keys.Length == 1)
                {
                    resultList.Add(keys[0].Trim().ToLower());
                    continue;
                }
                if (keys.Length != 2)
                    return null;
                // remove parenthese
                for (int i = 0; i < keys.Length; i++)
                {
                    string partten = "[a-z]";
                    keys[i] = Regex.Match(keys[i], partten).Value;
                }

                int indexOfLastKey = IsInKeysBL(keys[1], resultList);
                int indexOfFirstKey = IsInKeysBL(keys[0], resultList);
                if (indexOfFirstKey == -1 && indexOfLastKey == -1)
                {
                    resultList.Add(keys[0]);
                    resultList.Add(keys[1]);
                }
                else if (indexOfFirstKey == -1 && indexOfLastKey != -1)
                {
                    resultList.Insert(indexOfLastKey, keys[0]);
                }
                else if (indexOfFirstKey != -1 && indexOfLastKey == -1)
                {
                    if (indexOfFirstKey == resultList.Count - 1)
                    {
                        resultList.Add(keys[1]);
                    }
                    else
                        resultList.Insert(indexOfFirstKey + 1, keys[1]);
                }
                else if (indexOfFirstKey > indexOfLastKey)
                {
                    // exchange the place
                    resultList.Insert(indexOfLastKey, keys[0]);
                    resultList.RemoveAt(indexOfLastKey + 1);

                    resultList.Insert(indexOfFirstKey, keys[1]);
                    resultList.RemoveAt(indexOfFirstKey + 1);
                }
            }
            return resultList;
        }

        private static int IsInKeysBL(string value, ArrayList ar)
        {
            int result = -1;
            foreach (var item in ar)
            {
                if (item.ToString() == value)
                {
                    result = ar.IndexOf(item);
                    break;
                }
            }
            return result;
        }

        public static ArrayList GetKeysForDD(string input)
        {
            ArrayList resultList = new ArrayList();

            // remove only
            string _only = "only";
            int _indexOfOnly = input.ToLower().IndexOf(_only);

            string validString = input;
            if (_indexOfOnly != -1)
            {
                validString = input.ToLower().Substring(0, _indexOfOnly);
            }

            string[] keys = validString.ToLower().Split(new string[] { "and" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var key in keys)
            {
                if (string.IsNullOrWhiteSpace(key))
                    continue;
                // each
                string[] values = key.Split(new string[] { "or" }, StringSplitOptions.RemoveEmptyEntries);
                // remove parenthese
                for (int i = 0; i < values.Length; i++)
                {
                    //fix bug make it to match f13
                    //string partten = @"[a-z][1-9]";
                    string partten = @"[a-z]\d+";
                    values[i] = Regex.Match(values[i], partten).Value;
                }

                foreach (var item in values)
                {
                    if (!IsInKeysDD(item, resultList))
                    {
                        resultList.Add(item);
                        break;
                    }
                }
            }

            return resultList;

        }

        private static bool IsInKeysDD(string value, ArrayList ar)
        {
            string partten = @"[a-z]";
            foreach (var item in ar)
            {
                string key = Regex.Match(item.ToString(), partten).Value;
                string val = Regex.Match(value, partten).Value;
                if (key == val)
                    return true;
            }

            return false;
        }

        public static void GetKeysForAS(string input, out string[] clickKeys, out string[] selectKeys)
        {
            ArrayList resultList = new ArrayList();
            string validString = input;
            int _indesOfOnly = input.ToLower().IndexOf("only");
            if (_indesOfOnly != -1)
            {
                validString = input.ToLower().Substring(0, _indesOfOnly);
            }

            // split by and
            string[] keys = validString.ToLower().Split(new string[] { "and" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var key in keys)
            {
                if (string.IsNullOrWhiteSpace(key))
                    continue;
                string[] orAnswers = key.Split(new string[] { "or" }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var item in orAnswers)
                {
                    resultList.Add(item);
                    break;
                }
            }
            StringBuilder sbStr = new StringBuilder();
            foreach (var item in resultList)
            {
                sbStr.Append(item);
                sbStr.Append(",");
            }

            MatchCollection mc = Regex.Matches(sbStr.ToString(), "[a-z]");
            clickKeys = new string[mc.Count];
            for (int i = 0; i < mc.Count; i++)
            {
                clickKeys[i] = mc[i].Value;
            }

            mc = Regex.Matches(sbStr.ToString(), @"[a-z]\d+");
            selectKeys = new string[mc.Count];
            for (int i = 0; i < mc.Count; i++)
            {
                selectKeys[i] = mc[i].Value;
            }
        }
    }
}
