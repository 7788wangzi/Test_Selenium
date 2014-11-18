using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Se = OpenQA.Selenium;
using System.IO;
using System.Text.RegularExpressions;
// drag and drop
using OpenQA.Selenium.Interactions;
using System.Collections;
using UI = OpenQA.Selenium.Support.UI;

namespace Test_Selenium
{
    public class ElementHP
    {
        public static Se.IWebDriver GetWebDriver(Browser browser)
        {
            switch (browser)
            {
                case Browser.IE:
                    {
                        Se.IE.InternetExplorerOptions _IEOptions = new Se.IE.InternetExplorerOptions();
                        _IEOptions.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
                        _IEOptions.UnexpectedAlertBehavior = Se.IE.InternetExplorerUnexpectedAlertBehavior.Accept;
                        _IEOptions.ElementScrollBehavior = Se.IE.InternetExplorerElementScrollBehavior.Bottom;
                        return new Se.IE.InternetExplorerDriver(_IEOptions);
                    };
                case Browser.Firefox:
                    {
                        return new Se.Firefox.FirefoxDriver();
                    };
                case Browser.chrome:
                    {
                        return new Se.Chrome.ChromeDriver();
                    };
                default:
                    {
                        Se.IE.InternetExplorerOptions _IEOptions = new Se.IE.InternetExplorerOptions();
                        _IEOptions.IntroduceInstabilityByIgnoringProtectedModeSettings = true;
                        _IEOptions.UnexpectedAlertBehavior = Se.IE.InternetExplorerUnexpectedAlertBehavior.Accept;
                        _IEOptions.ElementScrollBehavior = Se.IE.InternetExplorerElementScrollBehavior.Bottom;
                        return new Se.IE.InternetExplorerDriver(_IEOptions);
                    };
            }
        }

        public static Se.IWebElement GetButton(string title, Se.IWebDriver driver)
        {
            Se.IWebElement button = null;
            driver.SwitchTo().DefaultContent();
            // page title
            var pageTitle = driver.Title;

            // all frames
            IList<Se.IWebElement> frames = driver.FindElements(Se.By.TagName("frame"));
            Se.IWebElement controlPanelFrame = null;
            foreach (var frame in frames)
            {
                if (frame.GetAttribute("name") == "ControlPanelFrame")
                {
                    controlPanelFrame = frame;
                    break;
                }
            }

            if (controlPanelFrame != null)
            {
                driver.SwitchTo().Frame(controlPanelFrame);
            }

            // all li elements in ControlPanelFrame
            IList<Se.IWebElement> liElements = driver.FindElements(Se.By.TagName("li"));
            foreach (var item in liElements)
            {
                // explicit the invisible buttons
                if (item.Displayed == false)
                    continue;
                if (item.GetAttribute("title").Trim().ToLower() == title.ToLower() && item.Text.Trim().ToLower() == title.ToLower())
                {
                    button = item;
                    break;
                }
            }
            return button;
        }

        public static string OpenTestsFolder(Se.IWebDriver driver)
        {
            driver.SwitchTo().DefaultContent();
            IList<Se.IWebElement> frames = driver.FindElements(Se.By.TagName("frame"));

            Se.IWebElement folderFrame = null;
            Se.IWebElement middleFrame = null;
            // click Tests node
            foreach (var item in frames)
            {
                if (item.GetAttribute("name") == "folders")
                {
                    folderFrame = item;
                    break;
                }
            }
            if (folderFrame != null)
            {
                driver.SwitchTo().Frame(folderFrame);
                string a = driver.Title;
                Se.IWebElement testsFolder = driver.FindElement(Se.By.Id("sd4"));
                if (testsFolder != null)
                {
                    testsFolder.Click();
                }
            }

            // for middle
            driver.SwitchTo().DefaultContent();
            frames = driver.FindElements(Se.By.TagName("frame"));
            foreach (var item in frames)
            {
                if (item.GetAttribute("name") == "middle")
                {
                    middleFrame = item;
                    break;
                }
            }

            if (middleFrame != null)
            {
                driver.SwitchTo().Frame(middleFrame);
                Se.IWebElement h1 = driver.FindElement(Se.By.XPath(@"//div[@id='header']/h1"));
                if (h1 != null)
                {
                    if (h1.Text.ToLower() == "tests")
                        return Message.Success;
                }
            }

            return Message.NodeNotFound;
        }

        public static string OpenExam(string exam, Se.IWebDriver driver)
        {
            driver.SwitchTo().DefaultContent();
            IList<Se.IWebElement> treeNodes = driver.FindElements(Se.By.TagName("frame"));

            // get middle frame
            Se.IWebElement middleFrame = null;
            foreach (var item in treeNodes)
            {
                if (item.GetAttribute("name") == "middle")
                {
                    middleFrame = item;
                    break;
                }
            }
            // Show all Active
            if (middleFrame != null)
            {
                driver.SwitchTo().Frame(middleFrame);
                Se.IWebElement showAllActive = driver.FindElement(Se.By.Id("lnkViewAllActive"));
                if (showAllActive != null)
                {
                    showAllActive.Click();
                }
            }

            driver.SwitchTo().DefaultContent();
            treeNodes = driver.FindElements(Se.By.TagName("frame"));
            // get middle frame again
            foreach (var item in treeNodes)
            {
                if (item.GetAttribute("name") == "middle")
                {
                    middleFrame = item;
                    break;
                }
            }
            driver.SwitchTo().Frame(middleFrame);
            Se.IWebElement examNode = driver.FindElement(Se.By.LinkText(exam));
            if (examNode != null)
            {
                examNode.Click();
                return Message.Success.ToString();
            }
            else
            {
                return Message.ExamNotFound.ToString();
            }
        }

        public static int FormsCount(Se.IWebDriver driver)
        {
            int count = 0;
            driver.SwitchTo().DefaultContent();
            IList<Se.IWebElement> treeNodes = driver.FindElements(Se.By.TagName("frame"));
            // get middle frame
            Se.IWebElement middleFrame = null;
            foreach (var item in treeNodes)
            {
                if (item.GetAttribute("name") == "middle")
                {
                    middleFrame = item;
                    break;
                }
            }
            // Show all Active
            if (middleFrame != null)
            {
                driver.SwitchTo().Frame(middleFrame);
                IList<Se.IWebElement> elements = driver.FindElements(Se.By.XPath(@"//table[@id='tblForms']/tbody/tr"));
                if (elements.Count > 1)
                    count = elements.Count - 1;
            }
            return count;
        }

        public static string OpenForm(int id, Se.IWebDriver driver)
        {
            driver.SwitchTo().DefaultContent();
            IList<Se.IWebElement> frames = driver.FindElements(Se.By.TagName("frame"));

            Se.IWebElement middleFrame = null;
            // one of the form
            foreach (var item in frames)
            {
                if (item.GetAttribute("name") == "middle")
                {
                    middleFrame = item;
                    break;
                }
            }

            if (middleFrame != null)
            {
                driver.SwitchTo().Frame(middleFrame);
                Se.IWebElement customButton = driver.FindElement(Se.By.XPath(@"//table[@id='tblForms']/tbody/tr[2]/td[last()]/ul/li[last()]/a"));

                if (customButton != null)
                {
                    customButton.Click();
                    CheckOptionsAndLaunch(driver);
                    return Message.Success.ToString();
                }
                return Message.FormNotFound.ToString();
            }
            return Message.FormNotFound.ToString();
        }

        private static void CheckOptionsAndLaunch(Se.IWebDriver driver)
        {
            driver.SwitchTo().DefaultContent();
            IList<Se.IWebElement> frames = driver.FindElements(Se.By.TagName("frame"));

            Se.IWebElement middleFrame = null;
            // one of the form
            foreach (var item in frames)
            {
                if (item.GetAttribute("name") == "middle")
                {
                    middleFrame = item;
                    break;
                }
            }

            if (middleFrame != null)
            {
                driver.SwitchTo().Frame(middleFrame);

                IList<Se.IWebElement> inputs = driver.FindElements(Se.By.ClassName("deliveryOpts"));
                int i = 0;
                foreach (var item in inputs)
                {
                    i++;
                    if (i == 10 || i == 20 || i == 21)
                    {
                        if (!item.Selected)
                            item.Click();
                    }
                }

                IList<Se.IWebElement> buttons = driver.FindElements(Se.By.XPath(@"//button"));
                foreach (var item in buttons)
                {
                    if (item.Text == "Preview")
                    {
                        item.Click();
                        break;
                    }
                }
            }
        }

        public static void DismissOptionPopup(Se.IWebDriver driver)
        {
            driver.SwitchTo().DefaultContent();
            IList<Se.IWebElement> frames = driver.FindElements(Se.By.TagName("frame"));

            Se.IWebElement middleFrame = null;
            // one of the form
            foreach (var item in frames)
            {
                if (item.GetAttribute("name") == "middle")
                {
                    middleFrame = item;
                    break;
                }
            }

            if (middleFrame != null)
            {
                driver.SwitchTo().Frame(middleFrame);

                IList<Se.IWebElement> buttons = driver.FindElements(Se.By.XPath(@"//button"));
                foreach (var item in buttons)
                {
                    if (item.Text == "Cancel")
                    {
                        item.Click();
                        break;
                    }
                }
            }
        }

        public static void DismissNDA(Se.IWebDriver driver)
        {
            IList<Se.IWebElement> frames = driver.FindElements(Se.By.TagName("frame"));
            Se.IWebElement frameDisplay = null;
            foreach (var frame in frames)
            {
                if (frame.GetAttribute("name") == "ElementDisplayFrame")
                {
                    frameDisplay = frame;
                    break;
                }
            }

            if (frameDisplay != null)
            {
                driver.SwitchTo().Frame(frameDisplay);

                // scroll to bottom
                Se.IWebElement ndaContainer = null;
                try
                {
                    ndaContainer = driver.FindElement(Se.By.Id("ndacontainer"));
                }
                catch { return; }
                if (ndaContainer == null)
                    return;
                string id = ndaContainer.GetAttribute("id");
                var js = "var q = document.getElementById('" + id + "').scrollTop=10000";
                //((Se.IE.InternetExplorerDriver)driver).ExecuteScript(js, null);
                switch (GlobalSets.CurrentBrower)
                {
                    case Browser.IE:
                        {
                            ((Se.IE.InternetExplorerDriver)driver).ExecuteScript(js, null);
                        }; break;
                    case Browser.chrome:
                        {
                            ((Se.Chrome.ChromeDriver)driver).ExecuteScript(js, null);
                        }; break;
                    case Browser.Firefox:
                        {
                            ((Se.Firefox.FirefoxDriver)driver).ExecuteScript(js, null);
                        }; break;
                    default:
                        {
                            ((Se.IE.InternetExplorerDriver)driver).ExecuteScript(js, null);
                        }; break;
                }

                Se.IWebElement yesButton = null;
                try
                {
                    yesButton = driver.FindElement(Se.By.XPath("//input[@name='I1']"));
                    //yesButton = driver.FindElement(Se.By.Id("chk11"));
                }
                catch { return; }

                //UI.WebDriverWait _wait = new UI.WebDriverWait(driver, TimeSpan.FromSeconds(50));
                //_wait.Until((d) => { return yesButton.Enabled == true; });
                if (yesButton != null && yesButton.Enabled)
                {
                    yesButton.Click();
                }

               // click yes span again
                Se.IWebElement yesSpan = null;
                yesSpan = driver.FindElement(Se.By.Id("nda-yes"));
                if (yesSpan != null)
                {
                    new Actions(driver).Click(yesSpan).DoubleClick(yesSpan).Build().Perform();
                }
            }
        }

        public static void SetAnswer_MC(Se.IWebDriver driver)
        {
            driver.SwitchTo().DefaultContent();

            IList<Se.IWebElement> frames = driver.FindElements(Se.By.TagName("frame"));
            Se.IWebElement elementDisplayFrame = null;
            foreach (var frame in frames)
            {
                if (frame.GetAttribute("name") == "ElementDisplayFrame")
                {
                    elementDisplayFrame = frame;
                    break;
                }
            }

            if (elementDisplayFrame != null)
            {
                // check if is MC item
                driver.SwitchTo().Frame(elementDisplayFrame);
                string itemType = driver.FindElements(Se.By.TagName("body"))[0].GetAttribute("class");
                if (itemType != ItemType.MC.ToString())
                    return;
                // get key
                Se.IWebElement keyElement = driver.FindElement(Se.By.XPath(@"//input[@name='Key1']"));
                //Se.IWebElement MaxSelect = driver.FindElement(Se.By.XPath(@"//input[@name='MaxSelectionAllowed']"));
                string[] keys = null;
                if (keyElement != null)
                {
                    string value = keyElement.GetAttribute("value");
                    keys = value.Split(new char[] { ',', ';' });
                }

                //bool isM1 = true;
                //// check if M1
                //if (MaxSelect != null)
                //{
                //    string value = MaxSelect.GetAttribute("value");
                //    if (value != "1")
                //        isM1 = false;
                //}

                //if (isM1)
                //{
                //    // set answer
                //    IList<Se.IWebElement> choices = driver.FindElements(Se.By.XPath(@"//input[@class='ITSMCOptionMarker']"));
                //    foreach (var choice in choices)
                //    {
                //        string choiceValue = choice.GetAttribute("value");
                //        if (IsInKeys(choiceValue, keys))
                //        {
                //            try
                //            {
                //                choice.Click();
                //            }
                //            catch { }
                //        }
                //    }
                //}
                //else
                //{
                IList<Se.IWebElement> clickableTables = driver.FindElements(Se.By.XPath(@"//table[@class='ITSMCOptionTable']"));
                foreach (var table in clickableTables)
                {
                    string keyseqValue = table.GetAttribute("keyseq");
                    if (IsInKeys(keyseqValue, keys))
                    {
                        try
                        {
                            table.Click();
                        }
                        catch { }
                    }
                }

                // re-check
                bool isChecked = true;
                foreach (var table in clickableTables)
                {
                    string keyseqValue = table.GetAttribute("keyseq");
                    if (IsInKeys(keyseqValue, keys))
                    {
                        if (!table.Selected)
                        {
                            isChecked = false;
                            break;
                        }
                    }
                }
                if (!isChecked)
                {
                    SetAnswer_MC(driver);
                }
                //}

            }
        }

        private static bool IsInKeys(string value, string[] keys)
        {
            foreach (var item in keys)
            {
                if (value.Trim().ToLower() == item.Trim().ToLower())
                    return true;
            }
            return false;
        }

        public static void SetAnswer_HA(Se.IWebDriver driver)
        {
            // div class name = "ITSHotArea"
            driver.SwitchTo().DefaultContent();

            IList<Se.IWebElement> frames = driver.FindElements(Se.By.TagName("frame"));
            Se.IWebElement elementDisplayFrame = null;
            foreach (var frame in frames)
            {
                if (frame.GetAttribute("name") == "ElementDisplayFrame")
                {
                    elementDisplayFrame = frame;
                    break;
                }
            }

            if (elementDisplayFrame != null)
            {
                // check if is HA item
                driver.SwitchTo().Frame(elementDisplayFrame);
                string itemType = driver.FindElements(Se.By.TagName("body"))[0].GetAttribute("class");
                if (itemType != ItemType.HA.ToString())
                    return;
                // get key
                Se.IWebElement keyElement = driver.FindElement(Se.By.XPath(@"//input[@name='Key1']"));
                if (keyElement != null)
                {
                    string value = keyElement.GetAttribute("value");

                    string[] keys = value.Split(new char[] { ',', ';' });

                    // set answer
                    IList<Se.IWebElement> choices = driver.FindElements(Se.By.XPath(@"//div[@class='ITSHotArea']"));
                    foreach (var choice in choices)
                    {
                        string choiceValue = choice.GetAttribute("value");
                        if (IsInKeys(choiceValue, keys))
                        {
                            try
                            {
                                choice.Click();
                            }
                            catch { }
                        }
                    }
                }
            }
        }

        // 
        public static void SetAnswer_DD(Se.IWebDriver driver)
        {
            // source div class ="ITSSourceAreaContainer" div class ="ITSDraggable ui-draggable ui-draggable-disabled"
            // target div class ="ITSWorkAreaContainer" ->  div class ="ITSDroppable ui-droppable"
            driver.SwitchTo().DefaultContent();

            IList<Se.IWebElement> frames = driver.FindElements(Se.By.TagName("frame"));
            Se.IWebElement elementDisplayFrame = null;
            foreach (var frame in frames)
            {
                if (frame.GetAttribute("name") == "ElementDisplayFrame")
                {
                    elementDisplayFrame = frame;
                    break;
                }
            }

            if (elementDisplayFrame != null)
            {
                // check if is DD item
                driver.SwitchTo().Frame(elementDisplayFrame);
                string itemType = driver.FindElements(Se.By.TagName("body"))[0].GetAttribute("class");
                if (itemType != ItemType.DD.ToString())
                    return;
                // get key
                Se.IWebElement keyElement = driver.FindElement(Se.By.XPath(@"//input[@name='Key1']"));

                // the current item "CurrentItems"
                Se.IWebElement itemElement = driver.FindElement(Se.By.XPath(@"//input[@name='CurrentItems']"));
                if (itemElement.GetAttribute("value") == "bbc_485")
                {
                    string breaks = true.ToString();
                }
                // sources and targets
                Se.IWebElement sourceContainer =null;
                sourceContainer = driver.FindElement(Se.By.XPath(@"//div[@class='ITSSourceAreaContainer']"));
                IList<Se.IWebElement> sources = null;
                if(sourceContainer!=null)
                {
                    sources = sourceContainer.FindElements(Se.By.XPath(@"//div[@class='ITSDraggable ui-draggable ui-draggable-handle']"));
                }

                Se.IWebElement targetContainer = null;
                targetContainer = driver.FindElement(Se.By.XPath(@"//div[@class='ITSWorkAreaContainer']"));
                IList<Se.IWebElement> targets = null;
                if (targetContainer != null)
                {
                    targets = targetContainer.FindElements(Se.By.XPath(@"//div[@class='ITSDroppable ui-droppable']"));
                }
                if (keyElement != null)
                {
                    string value = keyElement.GetAttribute("value");

                    //string partten = "[a-z|A-Z][1-9]";
                    //MatchCollection mc = Regex.Matches(value, partten);
                    //string[] keys = new string[mc.Count];
                    //for (int i = 0; i < mc.Count; i++)
                    //{
                    //    string key = mc[i].Value;
                    //    if (!IsInKeys(key, keys))
                    //        keys[i] = mc[i].Value;
                    //}

                    ArrayList keys = KeysHP.GetKeysForDD(value);

                    foreach (var answer in keys)
                    {
                        string key = answer.ToString();
                        if (!string.IsNullOrWhiteSpace(key))
                        {
                            string partten = "[a-y|A-Y]";
                            string targetID = Regex.Match(key, partten).Value;
                            //partten = "[1-9]";
                            partten = @"\d{1,2}";
                            string sourceID = Regex.Match(key, partten).Value;
                            if (!string.IsNullOrWhiteSpace(targetID) && !string.IsNullOrWhiteSpace(sourceID))
                            {
                                // get source and target 
                                Se.IWebElement source = null;
                                Se.IWebElement target = null;
                                foreach (var item in sources)
                                {
                                    if (item.GetAttribute("value").ToLower() == sourceID.ToLower())
                                    {
                                        source = item;
                                        break;
                                    }
                                }
                                foreach (var item in targets)
                                {
                                    if (item.GetAttribute("value").ToLower() == targetID.ToLower())
                                    {
                                        target = item;
                                        break;
                                    }
                                }

                                if (source != null && target != null)
                                {
                                    // drag and drop
                                    new Actions(driver).DragAndDrop(source, target).Perform();
                                }
                            }
                        }
                    }
                }
            }
        }

        // only work for Check box now
        public static void SetAnswer_AS(Se.IWebDriver driver)
        {
            driver.SwitchTo().DefaultContent();

            IList<Se.IWebElement> frames = driver.FindElements(Se.By.TagName("frame"));
            Se.IWebElement elementDisplayFrame = null;
            foreach (var frame in frames)
            {
                if (frame.GetAttribute("name") == "ElementDisplayFrame")
                {
                    elementDisplayFrame = frame;
                    break;
                }
            }

            if (elementDisplayFrame != null)
            {
                // check if is AS item
                driver.SwitchTo().Frame(elementDisplayFrame);
                string itemType = driver.FindElements(Se.By.TagName("body"))[0].GetAttribute("class");
                if (itemType != ItemType.AS.ToString())
                    return;
                // get key
                Se.IWebElement keyElement = driver.FindElement(Se.By.XPath(@"//input[@name='Key1']"));
                string[] clickableKeys = null;
                string[] selectableKeys = null;
                if (keyElement != null)
                {
                    string value = keyElement.GetAttribute("value");
                    KeysHP.GetKeysForAS(value, out clickableKeys, out selectableKeys);
                    //string partten = "[a-z|A-Z]";
                    //MatchCollection mc = Regex.Matches(value, partten);
                    //clickableKeys = new string[mc.Count];
                    //for (int i = 0; i < mc.Count; i++)
                    //{
                    //    string keyValue = mc[i].Value;
                    //    if (!IsInKeys(keyValue, clickableKeys))
                    //    {
                    //        clickableKeys[i] = keyValue;
                    //    }
                    //}
                }

                // div, get work area element
                Se.IWebElement workAreaElement = null;
                try
                {
                    workAreaElement = driver.FindElement(Se.By.XPath(@"//div[@class='ITSWorkAreaContainer']"));
                }
                catch { }
                if (workAreaElement == null)
                    return;

                if (clickableKeys.Length != 0)
                {

                    // get all check boxes
                    IList<Se.IWebElement> checkBoxes = workAreaElement.FindElements(Se.By.XPath(@"//input[@class='ITSCheckBox']"));
                    foreach (var item in checkBoxes)
                    {
                        item.Clear();
                        string value = item.GetAttribute("value").Trim().ToLower();
                        if (IsInKeys(value, clickableKeys))
                        {
                            try
                            {
                                item.Click();
                            }
                            catch { }
                        }
                    }

                    // for radio button
                    IList<Se.IWebElement> radioButtons = workAreaElement.FindElements(Se.By.XPath(@"//input[@class='ITSOptionBox']"));
                    foreach (var item in radioButtons)
                    {
                        string value = item.GetAttribute("value").Trim().ToLower();
                        if (IsInKeys(value, clickableKeys))
                        {
                            try
                            {
                                item.Click();
                            }
                            catch { }
                        }
                    }

                    // recheck for radio button
                    bool isChecked = true;
                    foreach (var item in radioButtons)
                    {
                        string value = item.GetAttribute("value").Trim().ToLower();
                        if (IsInKeys(value, clickableKeys))
                        {
                            if (!item.Selected)
                            {
                                isChecked = false;
                                break;
                            }
                        }
                    }
                    if (!isChecked)
                    {
                        SetAnswer_AS(driver);
                    }
                }

                if (selectableKeys.Length == 0)
                    return;
                // for select combo box, spin box
                foreach (var key in selectableKeys)
                {
                    string letter = Regex.Match(key.ToLower(), "[a-z]").Value;
                    string idOfComboBox = string.Format("itdComboBox{0}-1", letter).ToLower();
                    IList<Se.IWebElement> comboDiv = workAreaElement.FindElements(Se.By.XPath("//div[@role='combobox']"));//driver.FindElements(Se.By.XPath("//div[@class='jqx-dropdownlist-content']"));
                    foreach (var item in comboDiv)
                    {
                        if (item.GetAttribute("id").ToLower() == idOfComboBox)
                        {
                            item.Click();
                            //driver.Navigate().Refresh();
                            // all sources
                            IList<Se.IWebElement> sources = workAreaElement.FindElements(Se.By.XPath("//div[@class='jqx-listitem-element']/span"));
                            foreach (var span in sources)
                            {
                                if (span.Text.ToLower().StartsWith(key))
                                {
                                    new Actions(driver).DragAndDrop(span, item).Perform();
                                    break;
                                }
                            }
                        }
                    }

                    string idOfSpinBox = string.Format(@"itdUpDown{0}-1", letter).ToLower();
                    IList<Se.IWebElement> spinDiv = workAreaElement.FindElements(Se.By.XPath("//div[@role='spinbutton']"));
                    foreach (var item in spinDiv)
                    {
                        if (item.GetAttribute("id").ToLower() == idOfSpinBox)
                        {
                            Se.IWebElement input = item.FindElement(Se.By.TagName("input"));
                            string number = Regex.Match(key, @"\d+").Value;
                            input.SendKeys(number);
                        }
                    }

                }
            }
        }

        public static void SetAnswer_BL(Se.IWebDriver driver)
        {
            driver.SwitchTo().DefaultContent();
            IList<Se.IWebElement> frames = driver.FindElements(Se.By.TagName("frame"));
            Se.IWebElement elementDisplayFrame = null;
            foreach (var frame in frames)
            {
                if (frame.GetAttribute("name") == "ElementDisplayFrame")
                {
                    elementDisplayFrame = frame;
                    break;
                }
            }

            if (elementDisplayFrame != null)
            {
                // check if is BL item
                driver.SwitchTo().Frame(elementDisplayFrame);
                string itemType = driver.FindElements(Se.By.TagName("body"))[0].GetAttribute("class");
                if (itemType != ItemType.BL.ToString())
                    return;
                // get key
                Se.IWebElement keyElement = driver.FindElement(Se.By.XPath(@"//input[@name='Key1']"));
                if (keyElement != null)
                {
                    string value = keyElement.GetAttribute("value");
                    ArrayList keys = KeysHP.GetKeysForBL(value);
                    Se.IWebElement target = driver.FindElement(Se.By.XPath(@"//div[@class='ITSScrollingAnswerArea']/ol"));
                    IList<Se.IWebElement> sources = driver.FindElements(Se.By.XPath(@"//div[@class='ITSScrollingListArea']/ul/li"));
                    foreach (var item in keys)
                    {
                        string key = item.ToString().Trim();
                        foreach (var source in sources)
                        {
                            Se.IWebElement spanElement = null;
                            try
                            {
                                spanElement = source.FindElement(Se.By.TagName("span"));
                            }
                            catch { }
                            if (spanElement == null)
                                return;
                            string labelKey = spanElement.Text.Trim().ToLower();
                            if (labelKey.Trim().ToLower() == key)
                            {
                                // drag and drop, 
                                //new Actions(driver).DragAndDrop(source, target).Perform();

                                // press enter key
                                source.SendKeys(Se.Keys.Enter);
                                break;
                            }
                        }
                    }

                    // re-check
                    IList<Se.IWebElement> targets = driver.FindElements(Se.By.XPath(@"//div[@class='ITSScrollingAnswerArea']/ol/li"));
                    bool isSet = true;
                    if (targets.Count != keys.Count)
                        isSet = false;

                    if (!isSet)
                    {
                        // click reset button
                        Se.IWebElement reset = null;
                        reset = GetButton("Reset", driver);
                        if (reset != null)
                            reset.Click();
                        SetAnswer_BL(driver);
                    }
                }
            }
        }

        public static void GiveComment(Se.IWebDriver driver)
        {
            // last page, then start commenting
            Se.IWebElement commentButton = null;
            commentButton = ElementHP.GetButton("Start Commenting", driver);
            if (commentButton != null)
            {
                commentButton.Click();
            }

            // comment each question
            Se.IWebElement nextButton = null;
            nextButton = ElementHP.GetButton("Next", driver);
            while (nextButton != null)
            {
                // to do give comment

                try
                {
                    nextButton.Click();
                }
                catch { }
                nextButton = ElementHP.GetButton("Next", driver);
                // finish button for 
                if (nextButton == null)
                    nextButton = ElementHP.GetButton("Finished", driver);
            }
        }

        public static void FinishExam(Se.IWebDriver driver)
        {
            Se.IWebElement exitButton = null;
            exitButton = ElementHP.GetButton("Exit", driver);
            if (exitButton != null)
            {
                //
                try
                {
                    exitButton.Click();
                    // handle new alert window
                    DismissDialog(driver);
                }
                catch { }
                // in score report page
                Se.IWebElement nextButton = ElementHP.GetButton("Next", driver);
                if (nextButton != null)
                    nextButton.Click();

                driver.SwitchTo().DefaultContent();
                // take a screen shot   
                Se.Screenshot screenShot = null;
                switch (GlobalSets.CurrentBrower)
                {
                    case Browser.IE:
                        {
                            screenShot = ((Se.IE.InternetExplorerDriver)driver).GetScreenshot();
                        }; break;
                    case Browser.chrome:
                        {
                            screenShot = ((Se.Chrome.ChromeDriver)driver).GetScreenshot();
                        }; break;
                    case Browser.Firefox:
                        {
                            screenShot = ((Se.Firefox.FirefoxDriver)driver).GetScreenshot();
                        }; break;
                    default:
                        {
                            screenShot = ((Se.IE.InternetExplorerDriver)driver).GetScreenshot();
                        }; break;
                }

                string fileName = GlobalExam.TestName + "_" + DateTime.UtcNow.Ticks.ToString() + ".jpg";
                string saveDirectory = System.IO.Path.Combine(GlobalSets.ResultSavePath, GlobalExam.Name);
                if (!System.IO.Directory.Exists(saveDirectory))
                    Directory.CreateDirectory(saveDirectory);
                string fullName = Path.Combine(saveDirectory, fileName);
                screenShot.SaveAsFile(fullName, System.Drawing.Imaging.ImageFormat.Jpeg);
                //GlobalExam.Results = new List<string>();
                //GlobalExam.Results.Add(fullName);

                // Dont' end the browser otherwise, cannot go ahead
                Se.IWebElement endButton = ElementHP.GetButton("End", driver);
                if (endButton != null)
                {
                    endButton.Click();
                    try
                    {
                        driver.Close();
                    }
                    catch { }
                }
            }
        }

        // switch focus onto a new window, knowing the page title
        public static string GoToWindow(string title, ref Se.IWebDriver driver)
        {
            // get all window handles
            IList<string> handlers = driver.WindowHandles;
            foreach (var winHandler in handlers)
            {
                driver.SwitchTo().Window(winHandler);
                if (driver.Title == title)
                {
                    return Message.Success.ToString();
                }
                else
                {
                    driver.SwitchTo().DefaultContent();
                }
            }

            return Message.WindowNotFound.ToString();
        }

        // for popups in Exit
        public static void DismissDialog(Se.IWebDriver driver)
        {
            driver.SwitchTo().DefaultContent();
            IList<Se.IWebElement> frames = driver.FindElements(Se.By.TagName("frame"));
            Se.IWebElement frameDisplay = null;
            foreach (var frame in frames)
            {
                if (frame.GetAttribute("name") == "ElementDisplayFrame")
                {
                    frameDisplay = frame;
                }
            }

            if (frameDisplay != null)
            {
                driver.SwitchTo().Frame(frameDisplay);
                IList<Se.IWebElement> buttons = driver.FindElements(Se.By.XPath(@"//div[@class='dialogbuttons']/a"));
                foreach (var button in buttons)
                {
                    if (button.Text.Trim().ToLower() == "Yes".ToLower() || button.Text.Trim().ToLower() == "Ok".ToLower())
                    {
                        button.Click();
                        break;
                    }
                }
            }
        }
    }
}
