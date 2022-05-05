// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Security;

namespace Microsoft.Dynamics365.UIAutomation.Sample.Web
{
    [TestClass]
    public class CreateLead
    {
        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestMethod]
        public void WEBTestCreateNewLead()
        {
            using (var xrmBrowser = new Api.Browser(TestSettings.Options))
            {
                xrmBrowser.Navigate(_xrmUri);

                xrmBrowser.LoginPage.ADFSLoginAction(_username, _password);
                xrmBrowser.GuidedHelp.ClosePendingMails();
                xrmBrowser.GuidedHelp.CloseGuidedHelp();

                xrmBrowser.Navigation.OpenSubArea("Comercial", "Leads");
                
                xrmBrowser.CommandBar.ClickCommand("New");

                OptionSet pnet_documenttype = new OptionSet();
                pnet_documenttype.Name = "pnet_documenttype";
                pnet_documenttype.Value = "102610000";

                OptionSet pnet_gender = new OptionSet();
                pnet_gender.Name = "pnet_gender";
                pnet_gender.Value = "1";

                xrmBrowser.ThinkTime(2000);

                xrmBrowser.Entity.SetValue("firstname", "Martin");
                xrmBrowser.Entity.SetValue("lastname", "Lopez");
                xrmBrowser.Entity.SetValue(pnet_documenttype);
                xrmBrowser.Entity.SetValue(pnet_gender);

                xrmBrowser.Entity.SetValue("pnet_documentnumber", "29469000");

                
                //xrmBrowser.CommandBar.ClickCommand("Save & Close");
                xrmBrowser.ThinkTime(2000);
            }
        }
    }
}