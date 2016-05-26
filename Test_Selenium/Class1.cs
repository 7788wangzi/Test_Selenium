using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Se = OpenQA.Selenium;
using UI = OpenQA.Selenium.Support.UI;

namespace Test_Selenium
{
    /// <summary>
    ///  Used Selenium Driver
    /// </summary>
    public class Class1
    {
        Se.IWebDriver _WebDriver = null;
        UI.WebDriverWait _WebWait = null;

        public void Run(string exam, Browser browser, string savePath)
        {
            InitBrowser(browser, savePath);
            A();
            B();
            C(exam);
            D();
        }

        // init Internet driver
        private void InitBrowser(Browser browser, string savePath)
        {
            GlobalSets.CurrentBrower = browser;
            GlobalSets.ResultSavePath = savePath;
            switch (browser)
            {
                case Browser.IE:
                    {
                        _WebDriver = (Se.IE.InternetExplorerDriver)ElementHP.GetWebDriver(browser);
                    };break;
                case Browser.chrome:
                    {
                        _WebDriver = (Se.Chrome.ChromeDriver)ElementHP.GetWebDriver(browser);
                    }; break;
                case Browser.Firefox:
                    {
                        _WebDriver = (Se.Firefox.FirefoxDriver)ElementHP.GetWebDriver(browser);
                    };break;
                default:
                    {
                        _WebDriver = (Se.IE.InternetExplorerDriver)ElementHP.GetWebDriver(browser);
                    };break;
            }

            // webwait
            _WebWait = new UI.WebDriverWait(_WebDriver, TimeSpan.FromSeconds(50));
        }

        // open the webpage in explorer
        private void A()
        {
            _WebDriver.Navigate().GoToUrl(@"http://staging.programworkshop.com");
        }

        // fill in log on information
        private void B()
        {
            var title = _WebDriver.Title;
            _WebWait.Until((d) => { return d.Title.ToLower().StartsWith("login"); });
            _WebDriver.FindElement(Se.By.Name("uname")).SendKeys("qixue@microsoft.com");
            _WebDriver.FindElement(Se.By.Name("pass")).Clear();
            _WebDriver.FindElement(Se.By.Name("pass")).SendKeys("");
            _WebDriver.FindElement(Se.By.Name("btnlogin")).Click();
        }

        private void C(string exam)
        {
            
            _WebWait.Until((d) => { return d.Title.ToLower().StartsWith("program"); });
            var title = _WebDriver.Title;

            if (ElementHP.OpenTestsFolder(_WebDriver) != Message.Success.ToString())
            {
                // try one more time
                if((ElementHP.OpenTestsFolder(_WebDriver) != Message.Success.ToString()))
                    return;
            }
            
            // global variables
            GlobalExam.Name = exam;//"342-ENU-202";
            
            
            // open exam 342-ENU-202
            if (ElementHP.OpenExam(GlobalExam.Name, _WebDriver) != Message.Success.ToString())
                return;

            // get forms
            int formsCount = ElementHP.FormsCount(_WebDriver);
            GlobalExam.FormsCount = formsCount;

            // run for each form
            for (int i = 1; i <= formsCount; i++)
            {
                GlobalExam.TestName = "Form"+i;
                ElementHP.OpenForm(i, _WebDriver);

                // switch to Content Display window
                _WebWait.Until((d) => { return ElementHP.GoToWindow("Content Display", ref _WebDriver) == Message.Success.ToString(); });
                
                ElementHP.DismissNDA(_WebDriver);

                // new code
                Se.IWebElement nextButton = null;
                _WebWait.Until((d) => {
                    nextButton = ElementHP.GetButton("Next", _WebDriver);
                    return nextButton != null;
                });
                
                while (nextButton != null)
                {
                    try
                    {
                        nextButton.Click();
                        if (nextButton.Text.Trim().ToLower() == "Finished".ToLower())
                        {
                            try
                            {
                                ElementHP.DismissDialog(_WebDriver);
                            }
                            catch { }
                        }
                        // set answer for mc
                        ElementHP.SetAnswer_MC(_WebDriver);
                        ElementHP.SetAnswer_HA(_WebDriver);
                        ElementHP.SetAnswer_DD(_WebDriver);
                        ElementHP.SetAnswer_AS(_WebDriver);
                        ElementHP.SetAnswer_BL(_WebDriver);
                    }
                    catch { }
                    nextButton = ElementHP.GetButton("Next", _WebDriver);
                    // finish button for 
                    if (nextButton == null)
                        nextButton = ElementHP.GetButton("Finished", _WebDriver);
                }
                //ElementHP.GiveComment(_IEDriver);

                ElementHP.FinishExam(_WebDriver);

                // go back to Program Workshop window
                if (ElementHP.GoToWindow("Program Workshop", ref _WebDriver) == Message.WindowNotFound.ToString())
                    return;

                ElementHP.DismissOptionPopup(_WebDriver);
            }
            // open form A
            //ElementHP.OpenForm(1, _IEDriver);
        }

        private void D()
        {
            //IList<string> windows = _WebDriver.WindowHandles;
            //foreach (var item in windows)
            //{
            //    _WebDriver.SwitchTo().Window(item);
            //    _WebDriver.Close();
            //}

            _WebDriver.Quit();
        }
    }
}
