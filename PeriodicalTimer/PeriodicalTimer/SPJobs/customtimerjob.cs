using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using System.Collections.Specialized;
using Microsoft.SharePoint.Utilities;

namespace PeriodicalTimer.SPJobs
{
    class SPTImerJobClass : SPJobDefinition
    {
        public SPTImerJobClass() : base() { }

        public SPTImerJobClass(string jobName, SPService service)
            : base(jobName, service, null, SPJobLockType.None)
        {
            this.Title = "Periodical Check Timer";
        }

        public SPTImerJobClass(string jobName, SPWebApplication webapp)
            : base(jobName, webapp, null, SPJobLockType.ContentDatabase)
        {
            this.Title = "Periodical Check Timer";
        }

        public override void Execute(Guid targetInstanceId)
        {
            updateListPeriodical();
            updateListExtension();
        }

        public void updateListPeriodical()
        {
            var dateAndTime = DateTime.Now;
            var date = dateAndTime.Date;

            string refIDtemp = "";
            string MyEnd = "";
            string Notify_Period = "";
            int Period_Age = 0;

            SPWebApplication webApp = this.Parent as SPWebApplication;
            SPList ONTrack = webApp.Sites["engineering"].RootWeb.Lists["Online Leak Sealing (CAF-SUP-09)"];


            SPListItemCollection ONTrackItems = ONTrack.GetItems(new SPQuery()
            {
                Query = "<Where>" +
                "<And>" +
                "<Neq><FieldRef Name='End' /><Value Type=\"Text\">True</Value></Neq>" +
                "<IsNotNull><FieldRef Name='Notify_Period' /></IsNotNull>" +
                "</And>" +
                "</Where>",
            });

            foreach (SPListItem TrackItems in ONTrackItems)
            {
                try
                {
                    if (TrackItems["Notify_Period"] != null)
                        Notify_Period = TrackItems["Notify_Period"].ToString();

                    if (TrackItems["Period_Age"] != null)
                        Period_Age = Int32.Parse(TrackItems["Period_Age"].ToString());

                    if (TrackItems["ID"] != null)
                        refIDtemp = TrackItems["ID"].ToString();

                    if (TrackItems["End"] != null)
                        MyEnd = TrackItems["End"].ToString();

                    if (refIDtemp != "" && Notify_Period != "")
                    {
                        DateTime Notify = Convert.ToDateTime(Notify_Period);

                        if (CheckPeriod(refIDtemp, Notify, Period_Age))
                        {
                            TrackItems["Notify_Period"] = DateTime.Now;
                            TrackItems.Update();
                        }

                    }
                }
                catch (Exception e)
                {
                    errorLog("Error: " + e.ToString(), "");
                }

            }
        }

        public bool CheckPeriod(string RefID, DateTime NotifyPeriod, int PeroidAge)
        {
            DateTime NextDate;
            DateTime NextRecommend;

            string[] cat = new string[] {"AIMD", "AMT"};

            try
            {


                SPWebApplication webApp2 = this.Parent as SPWebApplication;
                SPList ONTrack2 = webApp2.Sites["engineering"].RootWeb.Lists["CAF-SUP-09 Online Leak Sealing Record"];
                
                bool upd = false;
                foreach (string catName in cat)
                {
                    SPListItemCollection ONTrackItems2 = ONTrack2.GetItems(new SPQuery()
                    {
                        Query = "<Where>" +
                                "<And>" +
                                "<Eq><FieldRef Name='Lookup_x0020_ID' /><Value Type=\"Text\">" + RefID + "</Value></Eq>" +
                                "<Eq><FieldRef Name='Category' /><Value Type=\"Text\">" + catName + "</Value></Eq>" +
                                "</And>" +
                                "</Where>",
                    });

                    DateTime next = NotifyPeriod.AddMonths(PeroidAge);
                    DateTime dateNow = DateTime.UtcNow;

                    NextDate = next;

                    if (ONTrackItems2.Count > 0)
                    {
                        SPListItem lt = ONTrackItems2[0];
                        NextRecommend = Convert.ToDateTime(lt["Next_x0020_Recommended_x0020_Ins"].ToString());                        
                    }
                    else
                    {
                        NextRecommend = NextDate;
                    }

                    if (dateNow > NextRecommend || ONTrackItems2.Count == 0)
                    {
                        SPListItem item = ONTrackItems2.Add();
                        item["Lookup_x0020_ID"] = RefID;
                        item["Next_x0020_Recommended_x0020_Ins"] = NextDate;
                        item["Category"] = catName;
                        item.Update();
                        upd = true;
                    }
                   
                }
                return upd;
            }
            catch (Exception e)
            {
                errorLog("SPListItemCollection Inspection  Error " + e.ToString(), "");
                return false;
            }
        }

        public void updateListExtension()
        {
            var dateAndTime = DateTime.Now;
            var date = dateAndTime.Date;
            
            string MyEnd = "";
            string Notify_Check = "";

            string typeOfRepair = "";
            string extension = "";
            DateTime expDate = DateTime.Now;
            DateTime extDate = DateTime.Now;

            SPWebApplication webApp = this.Parent as SPWebApplication;
            SPList ONTrack = webApp.Sites["engineering"].RootWeb.Lists["Online Leak Sealing (CAF-SUP-09)"];

            SPListItemCollection ONTrackItems = ONTrack.GetItems(new SPQuery()
            {
                Query = "<Where>" +
                "<And>" +
                    "<Neq><FieldRef Name='End' /><Value Type=\"Text\">True</Value></Neq>" +
                    "<IsNotNull><FieldRef Name='Notify_Check' /></IsNotNull>" +
                "</And>" +
                "</Where>",
            });

            foreach (SPListItem TrackItems in ONTrackItems)
            {
                try
                {
                    if (TrackItems["Notify_Check"] != null)
                        Notify_Check = TrackItems["Notify_Check"].ToString();

                    if (TrackItems["Type_x0020_of_x0020_Repair"] != null)
                        typeOfRepair = TrackItems["Type_x0020_of_x0020_Repair"].ToString();

                    if (TrackItems["Extension"] != null)
                        extension = TrackItems["Extension"].ToString();

                    if (TrackItems["End"] != null)
                        MyEnd = TrackItems["End"].ToString();

                    if (extension.Equals("Initiate") || extension.Equals("First"))
                    {
                        DateTime Notify = Convert.ToDateTime(Notify_Check);

                        SPWebApplication webApp2 = this.Parent as SPWebApplication;
                        SPList ONTrack2 = webApp2.Sites["engineering"].RootWeb.Lists["ExtensionTimer"];

                        SPListItemCollection ExtItem = ONTrack2.GetItems(new SPQuery()
                        {
                            Query = "<Where>" +
                            "<And>" +
                                "<Eq><FieldRef Name='Title' /><Value Type=\"Text\">" + typeOfRepair + "</Value></Eq>" +
                                "<Eq><FieldRef Name='ext' /><Value Type=\"Text\">" + extension + "</Value></Eq>" +
                            "</And>" +
                            "</Where>",
                        });

                        foreach (SPListItem ext in ExtItem)
                        {
                            int period = Int32.Parse(ext["Period"].ToString());
                            expDate = Notify.AddYears(period);
                            extDate = expDate.AddMonths(-1);
                        }

                        if (date >= extDate)
                        {
                            TrackItems["Workflow_x0020_Status"] = "Section 2 (i): Process Data Review Prior Executing Sealing Work by CRP";
                            TrackItems["IE_App"] = null;
                            TrackItems["Inspection_x0020_Approval_x0020_"] = null;
                            TrackItems["Tech_x0020_Approval_x0020_Date"] = null;
                            TrackItems["Static_App"] = null;
                            TrackItems["Safety_App"] = null;
                            TrackItems["OE_x0020_Approval_x0020_Date"] = null;
                            TrackItems["Sect3_x0020_HSE_x0020_Approval_x0"] = null;
                            TrackItems["TP_x0020_Inspection_x0020_Approv0"] = null;
                            TrackItems["TP_x0020_Technology_x0020_Approv0"] = null;
                            TrackItems["Sect3_x0020_Static_x002f_Piping_0"] = null;
                            TrackItems["Sect3_x0020_OE_x0020_Approval_x00"] = null;
                            TrackItems["Sect3_x0020_Other_x0020_Approval0"] = null;
                            TrackItems["Sect3_Other2_Approval"] = null;
                            TrackItems["Sect4_x0020_DE_x0020_Date1"] = null;
                            TrackItems["Sect4_x0020_DE_x0020_Date2"] = null;
                            TrackItems["Sect4_x0020_DE_x0020_Date3"] = null;
                            TrackItems["Sect4_x0020_DE_x0020_Date4"] = null;
                            TrackItems["Sect4_x0020_DE_x0020_Date5"] = null;
                            TrackItems["Sect4_x0020_DE_x0020_Date6"] = null;
                            TrackItems["Sect4_x0020_DE_x0020_Date7"] = null;
                            TrackItems["Sect4_x0020_DE_x0020_Date8"] = null;
                            TrackItems["Sect4_x0020_DE_x0020_Date9"] = null;
                            TrackItems["Assigned_x0020_To"] = TrackItems["CRP_x0020_Name"];
                            TrackItems["Current_x0020_Section"] = "Section 2";

                            if (extension.Equals("Initiate"))
                            {
                                TrackItems["Extension"] = "First";
                            }
                            else if (extension.Equals("First"))
                            {
                                TrackItems["Extension"] = "Second";
                            }
                            TrackItems["Notify_Check"] = DateTime.Now;
                            TrackItems.Update();
                            SendEmail(TrackItems["CRP_x0020_Name"].ToString(), TrackItems["CAF_x0020_No"].ToString()); 
                        }
                    }
                }
                catch (Exception e)
                {
                    errorLog("Error: " + e.ToString(), "");
                }

            }
        }

        public void SendEmail(string CRP, string CafNo)
        {
            try
            {
                using (SPSite oSPSite = SPContext.Current.Site)  //Site collection URL
                {
                    using (SPWeb oSPWeb = oSPSite.OpenWeb())  //Subsite URL
                    {
                        SPFieldLookupValue SingleValue = new SPFieldLookupValue(CRP);
                        int SPLookupID = SingleValue.LookupId;
                        SPUser spUser = new SPFieldUserValue(oSPWeb, Convert.ToInt32(SPLookupID), null).User;

                        StringDictionary headers = new StringDictionary();

                        headers.Add("from", "moc_mrcsb@petronas.com");
                        headers.Add("to", spUser.Email);
                        //headers.Add("bcc","SharePointAdmin@domain.com");
                        headers.Add("subject", "Extension of Online Leak Sealing CAF-SUP-09");
                        headers.Add("fAppendHtmlTag", "True"); //To enable HTML format

                        System.Text.StringBuilder strMessage = new System.Text.StringBuilder();
                        strMessage.Append("Greetings from Online MOC System,");
                        strMessage.Append("");
                        strMessage.Append("Dear " + spUser.Name + ",");
                        strMessage.Append("");
                        strMessage.Append("Please kindly review your extension for CAF Form - " + CafNo);
                        strMessage.Append("");
                        strMessage.Append("Regards,");
                        strMessage.Append("");
                        strMessage.Append("MOC Administrator.");

                        SPUtility.SendEmail(oSPWeb, headers, strMessage.ToString());

                    }
                }
            }
            catch { }
        }

        public void errorLog(string Desc, string Type)
        {
            SPWebApplication webapp = this.Parent as SPWebApplication;

            SPList tasklist = webapp.Sites["engineering"].RootWeb.Lists["JobList"];
            SPListItem newTask = tasklist.Items.Add();
            newTask["Title"] = "Periodical Log ";
            newTask["Description"] = Desc;
            newTask.Update();
        }

    }
}
