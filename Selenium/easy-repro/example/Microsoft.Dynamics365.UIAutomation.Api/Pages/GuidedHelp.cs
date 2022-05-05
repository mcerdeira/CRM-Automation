// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Dynamics365.UIAutomation.Browser;
using OpenQA.Selenium;
using System;

namespace Microsoft.Dynamics365.UIAutomation.Api
{
    /// <summary>
    /// Xrm Guided Help Page
    /// </summary>
    public class GuidedHelp
        : XrmPage
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="GuidedHelp"/> class.
        /// </summary>
        /// <param name="browser">The browser.</param>
        public GuidedHelp(InteractiveBrowser browser)
            : base(browser)
        {
            SwitchToDefault();
        }

        public void Dummy() 
        { 
        }

        public bool IsEnabled
        {
            get
            {
                bool isGuidedHelpEnabled = false;
                try
                {
                    bool.TryParse(
                        this.Browser.Driver.ExecuteScript("return Xrm.Internal.isFeatureEnabled('FCB.GuidedHelp') && Xrm.Internal.isGuidedHelpEnabledForUser();").ToString(),
                        out isGuidedHelpEnabled);
                }
                catch (System.NullReferenceException)
                {
                    // do nothing
                }


                return true; //isGuidedHelpEnabled;
            }
        }

        public BrowserCommandResult<bool> CloseDialog(int thinkTime = Constants.DefaultThinkTime)
        {
            this.Browser.ThinkTime(thinkTime);
            return this.Execute(GetOptions("Close Dialog"), driver =>
            {
                bool returnValue = false;

                if (driver.FindElement(By.Id("InlineDialog_Iframe")) != null)
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.Id("InlineDialog_Iframe")));
                    var but = driver.FindElement(By.Id("ok_id"));
                    but.Click();
                }

                return returnValue;
            });
        }

        public BrowserCommandResult<bool> ClosePendingMails(int thinkTime = Constants.DefaultThinkTime)
        {
            this.Browser.ThinkTime(thinkTime);
            return this.Execute(GetOptions("Close Pending Mails"), driver =>
            {
                bool returnValue = false;

                if (driver.FindElement(By.Id("InlineDialog1_Iframe")) != null)
                {
                    driver.SwitchTo().Frame(driver.FindElement(By.Id("InlineDialog1_Iframe")));
                    var but = driver.FindElement(By.Id("butBegin"));
                    but.Click();
                }

                return returnValue;
            });
        }

        /// <summary>
        /// Closes the Guided Help
        /// </summary>
        /// <example>xrmBrowser.GuidedHelp.CloseGuidedHelp();</example>
        public BrowserCommandResult<bool> CloseGuidedHelp(int thinkTime = Constants.DefaultThinkTime)
        {
            this.Browser.ThinkTime(thinkTime);
            return this.Execute(GetOptions("Close Guided Help"), driver =>
            {
                bool returnValue = false;

                driver.SwitchTo().Frame(driver.FindElement(By.Id("InlineDialog_Iframe")));


                var allMarsElements = driver
                        .FindElements(By.XPath(".//*"));

                foreach (var element in allMarsElements)
                {
                        
                    if (element.TagName == "img")
                    {
                        element.Click();
                        break;
                    }
                }
   
                return returnValue;
            });
        }
    }
}