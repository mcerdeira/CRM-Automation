using Microsoft.Dynamics365.UIAutomation.Api;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Security;

namespace Microsoft.Dynamics365.UIAutomation.Sample.Web.PROME
{
    [TestClass]
    public class Circuito1
    {
        private readonly SecureString _username = System.Configuration.ConfigurationManager.AppSettings["OnlineUsername"].ToSecureString();
        private readonly SecureString _password = System.Configuration.ConfigurationManager.AppSettings["OnlinePassword"].ToSecureString();
        private readonly Uri _xrmUri = new Uri(System.Configuration.ConfigurationManager.AppSettings["OnlineCrmUrl"].ToString());

        [TestMethod]
        public void WebCircuito1()
        {
            using (var xrmBrowser = new Api.Browser(TestSettings.Options))
            {
                #region Login
                xrmBrowser.Navigate(_xrmUri);
                xrmBrowser.LoginPage.ADFSLoginAction(_username, _password);
                xrmBrowser.GuidedHelp.ClosePendingMails();
                xrmBrowser.GuidedHelp.CloseGuidedHelp();
                #endregion

                #region Tomar datos estáticos
                // TODO: Tomar desde algún origen externo
                string contact_id = "9BB1D03E-C940-EC11-80D4-005056B20383";

                // TODO: Debería ser el usuario logueado
                string pnet_ejecutiveresponsible = "A1E1EA7E-A525-E511-80C2-00155D100F03";
                double pnet_estimatedamount = 140000;
                OptionSet pnet_svreporttype = new OptionSet();
                pnet_svreporttype.Name = "pnet_svreporttype";
                pnet_svreporttype.Value = "102610000";

                OptionSet pnet_svcredittype = new OptionSet();
                pnet_svcredittype.Name = "pnet_svcredittype";
                pnet_svcredittype.Value = "102610000";

                OptionSet pnet_svcreditsubtype = new OptionSet();
                pnet_svcreditsubtype.Name = "pnet_svcreditsubtype";
                pnet_svcreditsubtype.Value = "102610002";

                OptionSet pnet_process_evaluate = new OptionSet();
                pnet_process_evaluate.Name = "pnet_process_evaluate";
                pnet_process_evaluate.Value = "102610002";

                #endregion

                #region Antecedente comercial
                // Abrir un nuevo AC
                string commercial_url = string.Format("{0}main.aspx?etc=10048&extraqs=%3f_CreateFromId%3d%257b{1}%257d%26_CreateFromType%3d2%26_searchText%3d%26etc%3d10048%26parentLookupControlId%3dlookup_Commercial_Background_i&histKey=189917451&newWindow=true&pagetype=entityrecord"
                    , _xrmUri.AbsoluteUri, contact_id);
                xrmBrowser.Navigate(commercial_url);

                xrmBrowser.ThinkTime(1000);

                xrmBrowser.CommandBar.ClickCommand("Save"); // Se guarda con form de Zennon

                xrmBrowser.Entity.SelectForm("Information"); // Cambio de form

                xrmBrowser.ThinkTime(1000);

                xrmBrowser.Entity.SetValueJS("pnet_relatedcontact1", "null", true);

                xrmBrowser.Entity.SetValue("pnet_businesshistorybpbarefresh", true);

                xrmBrowser.LoginPage.CloseAlert();

                xrmBrowser.ThinkTime(3000);

                xrmBrowser.Entity.SaveSimple();

                xrmBrowser.Entity.SetValueJS("pnet_classification", "1", true);

                xrmBrowser.Entity.SaveSimple();

                Guid ac_guid = xrmBrowser.Entity.GetRecordGuid(); // Me guardo el guid del AC
                string ac_name = "AC"; //xrmBrowser.Entity.GetValue("pnet_code");
                #endregion  

                #region Entrvista Inicial
                // Abrir entrevista inicial
                string entrevista_url = string.Format("{0}main.aspx?etc=10015&extraqs=%3f_CreateFromId%3d%257b{1}%257d%26_CreateFromType%3d2%26etc%3d10015&histKey=947765298&newWindow=true&pagetype=entityrecord"
                    , _xrmUri.AbsoluteUri, contact_id);
                xrmBrowser.Navigate(entrevista_url);
                xrmBrowser.ThinkTime(1000);

                xrmBrowser.Entity.SetValueJSLookUp("pnet_commercialbackground", ac_guid.ToString(), ac_name, "pnet_commercialbackground", true);

                xrmBrowser.CommandBar.ClickCommand("Save");

                xrmBrowser.CommandBar.ClickCommand("Resultado EI");

                xrmBrowser.ThinkTime(2000);

                xrmBrowser.GuidedHelp.CloseDialog();

                xrmBrowser.ThinkTime(1000);

                #endregion

                #region Abrir el contacto para ir a la VT
                string contact_url = string.Format(
                    "{0}main.aspx?etc=2&extraqs=&histKey=47240210&id={1}&newWindow=true&pagetype=entityrecord&sitemappath=SFA%7cCustomers%7cnav_conts#51602170"
                    , _xrmUri.AbsoluteUri, contact_id);
                xrmBrowser.Navigate(contact_url);

                xrmBrowser.ThinkTime(5000);

                // Grilla de VTS
                int count = 5;
                while (true)
                {
                    List<GridItem> vts = xrmBrowser.Entity.GetSubGridItems("SITEVISITS");
                    if (vts.Count > 0)
                    {
                        foreach (var vt in vts)
                        {
                            Guid first = vt.Id;
                            xrmBrowser.Entity.OpenEntity("pnet_sitevisit", first);
                            break;
                        }
                        break;
                    }
                    else
                    {
                        count--;
                        if (count > 0)
                        {
                            xrmBrowser.ThinkTime(200);
                        }
                        else
                        {
                            Assert.Fail("No hay VT!!");
                            break;
                        }
                    }
                }
                #endregion

                #region Visita en Terreno
                xrmBrowser.ThinkTime(1000);

                xrmBrowser.Entity.SetValueJSLookUp("pnet_ejecutiveresponsible", pnet_ejecutiveresponsible, "Ejecutivo", "systemuser", true);
                xrmBrowser.Entity.SetValue("pnet_estimatedamount", pnet_estimatedamount.ToString());
                xrmBrowser.Entity.SetValue(pnet_svreporttype);
                xrmBrowser.Entity.SetValue(pnet_svcredittype);
                xrmBrowser.Entity.SetValue(pnet_svcreditsubtype);
                xrmBrowser.Entity.SetValue(pnet_process_evaluate);

                xrmBrowser.CommandBar.ClickCommand("Resultado VT");

                xrmBrowser.ThinkTime(2000);

                xrmBrowser.GuidedHelp.CloseDialog();
                #endregion


                #region Abrir el contacto para ir a la Solicitud
                xrmBrowser.Navigate(contact_url);

                xrmBrowser.ThinkTime(7000);
                // Grilla de Solicitudes
                count = 5;
                while (true)
                {
                    List<GridItem> ops = xrmBrowser.Entity.GetSubGridItems("Request");
                    if (ops.Count > 0)
                    {
                        foreach (var op in ops)
                        {
                            Guid fop = op.Id;
                            xrmBrowser.Entity.OpenEntity("opportunity", fop);
                            break;
                        }
                        break;
                    }
                    else
                    {
                        count--;
                        if (count > 0)
                        {
                            xrmBrowser.ThinkTime(200);
                        }
                        else
                        {
                            Assert.Fail("No hay Solicitud!!");
                            break;
                        }
                    }
                }
                xrmBrowser.ThinkTime(5000);
                #endregion
            }
        }
    }
}
