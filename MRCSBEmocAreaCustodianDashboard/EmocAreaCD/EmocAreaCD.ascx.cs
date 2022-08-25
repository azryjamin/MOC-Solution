using System;
using System.Xml;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Drawing;
using System.ComponentModel;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
//using System.Web.UI.DataVisualization.Charting;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Utilities;
using Microsoft.SharePoint.WebControls;
using System.Globalization;
using System.IO;

namespace MRCSBEmocAreaCustodianDashboard.VisualWebPart1
{
    [ToolboxItemAttribute(false)]
    public partial class VisualWebPart1 : WebPart
    {
        // Uncomment the following SecurityPermission attribute only when doing Performance Profiling using
        // the Instrumentation method, and then remove the SecurityPermission attribute when the code is ready
        // for production. Because the SecurityPermission attribute bypasses the security check for callers of
        // your constructor, it's not recommended for production purposes.
        // [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.Assert, UnmanagedCode = true)]
        public VisualWebPart1()
        {
        }

        protected override void OnInit(EventArgs e)
        {
            base.OnInit(e);
            InitializeControl();
        }

        public class getCAFDetails
        {
            public string CAFNo { get; set; }
            public string CAFType { get; set; }
            public string Title { get; set; }
            public string CRP { get; set; }
            public string CurrentStatus { get; set; }
            public string InitiationDate { get; set; }
            public string ConstructionDate { get; set; }
            public string StartUpDate { get; set; }
            public string SubmissionDate { get; set; }
            public string Remarks { get; set; }
        }

        public class getOpenAreaCAF
        {
            public string area { get; set; }
            public string count { get; set; }
        }

        public class getCloseAreaCAF
        {
            public string area { get; set; }
            public string count { get; set; }
        }

        public class getOverdueAreaCAF
        {
            public string area { get; set; }
            public string count { get; set; }
        }

        public class getNCAreaCAF
        {
            public string area { get; set; }
            public string count { get; set; }
        }

        public class getGTOverdue2weeks
        {
            public string area { get; set; }
            public string count { get; set; }
        }

        public class getGTNC2weeks
        {
            public string area { get; set; }
            public string count { get; set; }
        }

        public class getGTOverdue1mth
        {
            public string area { get; set; }
            public string count { get; set; }
        }

        public class getGTNC1mth
        {
            public string area { get; set; }
            public string count { get; set; }
        }

        public class getExtOLS
        {
            public string area { get; set; }
            public string count { get; set; }
        }

        public class getExtOLS1M
        {
            public string area { get; set; }
            public string count { get; set; }
        }
        protected void Page_Load(object sender, EventArgs e)
        {
            // on page load
            string URLwebmoc = "";
            URLwebmoc = SPContext.Current.Web.Url;

            lblDate.Text = DateTime.Now.Day.ToString() + "/" + DateTime.Now.Month.ToString() + "/" + DateTime.Now.Year.ToString();

            /*SPSite oSite;
            try
            {
                oSite = new SPSite(URLwebmoc);
            }
            catch
            {
                URLwebmoc = SPContext.Current.Web.Url;
                oSite = new SPSite(URLwebmoc);
            }*/

            using (SPSite oSite = new SPSite(URLwebmoc))
            {
                using (SPWeb oWeb = oSite.OpenWeb())
                {
                    /* Department list */
                    //SPList oListDept = oWeb.Lists["Department"];

                    /* CAF List */
                    SPList oList04 = oWeb.Lists["Standard Form (CAF-SUP-04)"];
                    SPList oList04A = oWeb.Lists["Non Process Area Change (CAF-SUP-04A)"];
                    SPList oList04B = oWeb.Lists["Transmitter Re-Ranging (CAF-SUP-04B)"];
                    SPList oList04C = oWeb.Lists["DCS Minor Change (CAF-SUP-04C)"];
                    SPList oList04D = oWeb.Lists["DCS Non Critical Alarm Parameter Change (CAF-SUP-04D)"];
                    SPList oList04E = oWeb.Lists["SCE Bypass (CAF-SUP-04E)"];
                    SPList oList04G = oWeb.Lists["Product Change (CAF-SUP-04G)"];
                    SPList oList04J = oWeb.Lists["Process Control Network Change (for Level 3 and above) (CAF-SUP-04J)"];
                    SPList oList09 = oWeb.Lists["Online Leak Sealing (CAF-SUP-09)"];
                    SPList oList04L = oWeb.Lists["Rejuvenate Change (CAF-SUP-04L)"];
                    //CAFSUP04L to be added
                    //SPList oList04L= oWeb.Lists["Total Status Report"];

                    //SPListItemCollection itemDept = oListDept.GetItems("Title");

                    SPListItemCollection item04 = oList04.GetItems("ID", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF"); //Items //nacb_090420_override_remarks
                    SPListItemCollection item04A = oList04A.GetItems("ID", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF"); //.GetItems("CAF_x0020_No", "CAF_x0020_Name", "Title", "Date_x0020_Initiated", "Est_x002e__x0020_Start_x0020_Up_", "Modified", "MOC_x0020_Status", "Area");
                    SPListItemCollection item04B = oList04B.GetItems("ID", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF");
                    SPListItemCollection item04C = oList04C.GetItems("ID", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF");
                    SPListItemCollection item04D = oList04D.GetItems("ID", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF");
                    SPListItemCollection item04E = oList04E.GetItems("ID", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "ByPass_x0020_Date");
                    SPListItemCollection item04G = oList04G.GetItems("ID", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF");
                    SPListItemCollection item04J = oList04J.GetItems("ID", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF");
                    SPListItemCollection item09 = oList09.GetItems("ID", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Type_x0020_of_x0020_Repair", "Extension", "Extension_Date", "Override_x0020_Remarks", "Actual_x0020_Date", "Sect3_HOP_App", "Sect4A_x0020_AO_x0020_Approval_x");
                    SPListItemCollection item04L = oList04L.GetItems("ID", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF");
                    Logger("Area AA1");
                    //DataTable dt04;//not initialized, will be null

                    DataTable dt04 = item04.GetDataTable();
                    DataTable dt04A = item04A.GetDataTable();
                    DataTable dt04B = item04B.GetDataTable();
                    DataTable dt04C = item04C.GetDataTable();
                    DataTable dt04D = item04D.GetDataTable();
                    DataTable dt04E = item04E.GetDataTable();
                    DataTable dt04G = item04G.GetDataTable();
                    DataTable dt04J = item04J.GetDataTable();
                    DataTable dt09 = item09.GetDataTable();
                    DataTable dt04L = item04L.GetDataTable();

                    #region Declare DataTable

                    if (item04.Count == 0)
                    {
                        dt04 = new DataTable();

                        dt04.Columns.Add("MOC_x0020_Status");
                        dt04.Columns.Add("Area");
                        dt04.Columns.Add("DATA_x0020_TYPE");
                        dt04.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04.Columns.Add("Current_x0020_Section");
                        dt04.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A2 04");

                    if (item04A.Count == 0)
                    {
                        dt04A = new DataTable();

                        dt04A.Columns.Add("MOC_x0020_Status");
                        dt04A.Columns.Add("Area");
                        dt04A.Columns.Add("DATA_x0020_TYPE");
                        dt04A.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04A.Columns.Add("Current_x0020_Section");
                        dt04A.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A2 04A");

                    if (item04B.Count == 0)
                    {
                        dt04B = new DataTable();

                        dt04B.Columns.Add("MOC_x0020_Status");
                        dt04B.Columns.Add("Area");
                        dt04B.Columns.Add("DATA_x0020_TYPE");
                        dt04B.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04B.Columns.Add("Current_x0020_Section");
                        dt04B.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A2 04B");

                    if (item04C.Count == 0)
                    {
                        dt04C = new DataTable();

                        dt04C.Columns.Add("MOC_x0020_Status");
                        dt04C.Columns.Add("Area");
                        dt04C.Columns.Add("DATA_x0020_TYPE");
                        dt04C.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04C.Columns.Add("Current_x0020_Section");
                        dt04C.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A2 04C");

                    if (item04D.Count == 0)
                    {
                        dt04D = new DataTable();

                        dt04D.Columns.Add("MOC_x0020_Status");
                        dt04D.Columns.Add("Area");
                        dt04D.Columns.Add("DATA_x0020_TYPE");
                        dt04D.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04D.Columns.Add("Current_x0020_Section");
                        dt04D.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A2 04D");

                    if (item04E.Count == 0)
                    {
                        dt04E = new DataTable();

                        dt04E.Columns.Add("MOC_x0020_Status");
                        dt04E.Columns.Add("Area");
                        dt04E.Columns.Add("DATA_x0020_TYPE");
                        dt04E.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04E.Columns.Add("Current_x0020_Section");
                        dt04E.Columns.Add("ByPass_x0020_Date");
                    }
                    Logger("Area A2 04E");

                    if (item04G.Count == 0)
                    {
                        dt04G = new DataTable();

                        dt04G.Columns.Add("MOC_x0020_Status");
                        dt04G.Columns.Add("Area");
                        dt04G.Columns.Add("DATA_x0020_TYPE");
                        dt04G.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04G.Columns.Add("Current_x0020_Section");
                        dt04G.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A2 04G");

                    if (item04J.Count == 0)
                    {
                        dt04J = new DataTable();

                        dt04J.Columns.Add("MOC_x0020_Status");
                        dt04J.Columns.Add("Area");
                        dt04J.Columns.Add("DATA_x0020_TYPE");
                        dt04J.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04J.Columns.Add("Current_x0020_Section");
                        dt04J.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A2 04J");

                    if (item09.Count == 0)
                    {
                        dt09 = new DataTable();

                        dt09.Columns.Add("MOC_x0020_Status");
                        dt09.Columns.Add("Area");
                        dt09.Columns.Add("DATA_x0020_TYPE");
                        dt09.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt09.Columns.Add("Current_x0020_Section");
                        dt09.Columns.Add("Type_x0020_of_x0020_Repair");
                        dt09.Columns.Add("Extension");
                        dt09.Columns.Add("Extension_Date");
                        dt09.Columns.Add("Actual_x0020_Date");
                        dt09.Columns.Add("Sect3_HOP_App");
                        dt09.Columns.Add("Sect4A_x0020_AO_x0020_Approval_x");
                    }
                    Logger("Area A2 09");

                    if (item04L.Count == 0)
                    {
                        dt04L = new DataTable();

                        dt04L.Columns.Add("MOC_x0020_Status");
                        dt04L.Columns.Add("Area");
                        dt04L.Columns.Add("DATA_x0020_TYPE");
                        dt04L.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04L.Columns.Add("Current_x0020_Section");
                        dt04.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A2 04L");

                    #endregion

                    ListViewCAFDetails.Visible = false;
                    btnBack.Visible = false;

                    DataTable tableEmocCAF;
                    // DataView tableEmocCAFView;

                    var today = DateTime.Today;

                    try
                    {

                        #region query for Open CAF 1
                        var getOpenCAF =
                                   (from item in dt04.Rows.Cast<DataRow>()
                                    where item["MOC_x0020_Status"].ToString().Equals("Open") && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                    && item["Current_x0020_Section"].ToString() != "Section 5"
                                    select item)
                                   .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                           where item2["MOC_x0020_Status"].ToString().Equals("Open") && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                           && item2["Current_x0020_Section"].ToString() != "Section 5"
                                           select item2)
                                   .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                           where item3["MOC_x0020_Status"].ToString().Equals("Open") && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                           && item3["Current_x0020_Section"].ToString() != "Section 5"
                                           select item3)
                                   .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                           where item4["MOC_x0020_Status"].ToString().Equals("Open") && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item4["Current_x0020_Section"].ToString() != "Section 5"
                                           select item4)
                                   .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                           where item5["MOC_x0020_Status"].ToString().Equals("Open") && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item5["Current_x0020_Section"].ToString() != "Section 5"
                                           select item5)
                                   .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                           where item6["MOC_x0020_Status"].ToString().Equals("Open") && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item6["Current_x0020_Section"].ToString() != "Section 5"
                                           select item6)
                                   .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                           where item7["MOC_x0020_Status"].ToString().Equals("Open") && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item7["Current_x0020_Section"].ToString() != "Section 5"
                                           select item7)
                                   .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                           where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item8["Current_x0020_Section"].ToString() != "Section 5"
                                           select item8)
                                   .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                           where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString() != "Section 5"
                                           select item9)
                            .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                    where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                     && item8["Current_x0020_Section"].ToString() != "Section 5"
                                    select item8)
                                   .GroupBy(s => new { Area = s["Area"] })
                                   .OrderBy(s => s.Key.Area)
                                   .Select(s => new
                                   {
                                       areaGroup = s.Key.Area,
                                       Count = s.Count()
                                   });
                        Logger("Area A4");
                        #endregion

                        #region query for Open CAF 2
                        var getOpenCAF2 =
                                   (from item in dt04.Rows.Cast<DataRow>()
                                    where item["MOC_x0020_Status"].ToString().Equals("Open") && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                    && item["Current_x0020_Section"].ToString() == "Section 5"
                                    select item)
                                   .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                           where item2["MOC_x0020_Status"].ToString().Equals("Open") && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                           && item2["Current_x0020_Section"].ToString() == "Section 5"
                                           select item2)
                                   .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                           where item3["MOC_x0020_Status"].ToString().Equals("Open") && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                           && item3["Current_x0020_Section"].ToString() == "Section 5"
                                           select item3)
                                   .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                           where item4["MOC_x0020_Status"].ToString().Equals("Open") && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item4["Current_x0020_Section"].ToString() == "Section 5"
                                           select item4)
                                   .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                           where item5["MOC_x0020_Status"].ToString().Equals("Open") && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item5["Current_x0020_Section"].ToString() == "Section 5"
                                           select item5)
                                   .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                           where item6["MOC_x0020_Status"].ToString().Equals("Open") && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item6["Current_x0020_Section"].ToString() == "Section 5"
                                           select item6)
                                   .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                           where item7["MOC_x0020_Status"].ToString().Equals("Open") && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item7["Current_x0020_Section"].ToString() == "Section 5"
                                           select item7)
                                   .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                           where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item8["Current_x0020_Section"].ToString() == "Section 5"
                                           select item8)
                                   .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                           where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString() == "Section 5"
                                           select item9)
                            .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                    where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                     && item8["Current_x0020_Section"].ToString() == "Section 5"
                                    select item8)
                                   .GroupBy(s => new { Area = s["Area"] })
                                   .OrderBy(s => s.Key.Area)
                                   .Select(s => new
                                   {
                                       areaGroup = s.Key.Area,
                                       Count = s.Count()
                                   });

                        #endregion

                        #region query for Close CAF
                        var getCloseCAF =
                                  (from item in dt04.Rows.Cast<DataRow>()
                                   where item["MOC_x0020_Status"].ToString().Equals("Closed") && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                   select item)
                                 .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                         where item2["MOC_x0020_Status"].ToString().Equals("Closed") && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                         select item2)
                                   .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                           where item3["MOC_x0020_Status"].ToString().Equals("Closed") && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                           select item3)
                                   .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                           where item4["MOC_x0020_Status"].ToString().Equals("Closed") && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                           select item4)
                                   .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                           where item5["MOC_x0020_Status"].ToString().Equals("Closed") && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                           select item5)
                                   .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                           where item6["MOC_x0020_Status"].ToString().Equals("Closed") && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                           select item6)
                                   .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                           where item7["MOC_x0020_Status"].ToString().Equals("Closed") && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                           select item7)
                                   .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                           where item8["MOC_x0020_Status"].ToString().Equals("Closed") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                           select item8)
                                   .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                           where item9["MOC_x0020_Status"].ToString().Equals("Closed") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                           select item9)
                            .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                    where item8["MOC_x0020_Status"].ToString().Equals("Closed") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                    select item8)
                                   .GroupBy(s => new { Area = s["Area"] })
                                   .OrderBy(s => s.Key.Area)
                                   .Select(s => new
                                   {
                                       areaGroup = s.Key.Area,
                                       Count = s.Count()
                                   });
                        Logger("Area A5");
                        #endregion

                        #region query for Overdue CAF
                        var getOverdueCAF = (from item in dt04.Rows.Cast<DataRow>()
                                             where item["MOC_x0020_Status"].ToString().Equals("Open") && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                             && ((item["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                             || (item["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                             select item)
                                   .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                           where item2["MOC_x0020_Status"].ToString().Equals("Open") && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && ((item2["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                            || (item2["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                           select item2)
                                   .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                           where item3["MOC_x0020_Status"].ToString().Equals("Open") && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && ((item3["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                            || (item3["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                           select item3)
                                    .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                            where item4["MOC_x0020_Status"].ToString().Equals("Open") && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                             && ((item4["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                             || (item4["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                            select item4)
                                    .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                            where item5["MOC_x0020_Status"].ToString().Equals("Open") && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                             && ((item5["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                             || (item5["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                            select item5)
                                    .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                            where item6["MOC_x0020_Status"].ToString().Equals("Open") && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                             && today > (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddMonths(6)
                                            select item6)
                                    .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                            where item7["MOC_x0020_Status"].ToString().Equals("Open") && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                             && ((item7["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                             || (item7["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                            select item7)
                                    .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                            where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                             && ((item8["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                             || (item8["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                            select item8)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "Initiate")
                                            && today > (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "First")
                                            && today > (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "First")
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today > (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "Second")
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today > (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Extension"].ToString() == "Initiate"
                                            && today > Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Extension"].ToString() == "First"
                                            && today > Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Extension"].ToString() == "First"
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today > Convert.ToDateTime(item9["Extension_Date"])
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Extension"].ToString() == "Second"
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today > Convert.ToDateTime(item9["Extension_Date"])
                                            select item9)
                            .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                    where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                     && ((item8["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                     || (item8["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                    select item8)
                                   .GroupBy(s => new { Area = s["Area"] })
                                   .OrderBy(s => s.Key.Area)
                                   .Select(s => new
                                   {
                                       areaGroup = s.Key.Area,
                                       Count = s.Count()
                                   });
                        Logger("Area A3 OD");
                        #endregion

                        #region query for NC
                        var getNCCAF = (from item in dt04.Rows.Cast<DataRow>()
                                        where item["MOC_x0020_Status"].ToString().Equals("Open") && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                        && today > Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])
                                        && item["Current_x0020_Section"].ToString() != "Section 5"
                                        select item)
                                   .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                           where item2["MOC_x0020_Status"].ToString().Equals("Open") && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today > Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])
                                            && item2["Current_x0020_Section"].ToString() != "Section 5"
                                           select item2)
                                   .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                           where item3["MOC_x0020_Status"].ToString().Equals("Open") && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today > Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])
                                            && item3["Current_x0020_Section"].ToString() != "Section 5"
                                           select item3)
                                   .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                           where item4["MOC_x0020_Status"].ToString().Equals("Open") && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today > Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])
                                            && item4["Current_x0020_Section"].ToString() != "Section 5"
                                           select item4)
                                   .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                           where item5["MOC_x0020_Status"].ToString().Equals("Open") && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today > Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])
                                            && item5["Current_x0020_Section"].ToString() != "Section 5"
                                           select item5)
                                   .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                           where item6["MOC_x0020_Status"].ToString().Equals("Open") && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today > Convert.ToDateTime(item6["ByPass_x0020_Date"]).AddDays(90)
                                            && item6["Current_x0020_Section"].ToString() != "Section 5"
                                           select item6)
                                   .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                           where item7["MOC_x0020_Status"].ToString().Equals("Open") && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today > Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])
                                            && item7["Current_x0020_Section"].ToString() != "Section 5"
                                           select item7)
                                   .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                           where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today > Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])
                                            && item8["Current_x0020_Section"].ToString() != "Section 5"
                                           select item8)
                                   .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                           where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                           && item9["Extension"].ToString() == ""
                                           && (item9["Current_x0020_Section"].ToString().Equals("Section 1")
                                           || item9["Current_x0020_Section"].ToString().Equals("Section 2")
                                           || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                           && today > Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddDays(14)
                                           select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Extension"].ToString() == ""
                                            && item9["Sect4A_x0020_AO_x0020_Approval_x"].ToString().Equals("")
                                            && !item9["Sect3_HOP_App"].ToString().Equals("")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 1")
                                            || item9["Current_x0020_Section"].ToString().Equals("Section 2")
                                            || item9["Current_x0020_Section"].ToString().Equals("Section 3")
                                            || item9["Current_x0020_Section"].ToString().Equals("Section 4"))
                                            && today > Convert.ToDateTime(item9["Sect3_HOP_App"]).AddDays(21)
                                            select item9)
                            .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                    where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                     && today > Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])
                                     && item8["Current_x0020_Section"].ToString() != "Section 5"
                                    select item8)
                                  .GroupBy(s => new { Area = s["Area"] })
                                  .OrderBy(s => s.Key.Area)
                                  .Select(s => new
                                  {
                                      areaGroup = s.Key.Area,
                                      Count = s.Count()
                                  });
                        Logger("Area A3 NC");
                        #endregion

                        #region query for Going to Overdue 2 weeks
                        var getGoingToOverdue2weekCAF = (from item in dt04.Rows.Cast<DataRow>()
                                                         where item["MOC_x0020_Status"].ToString().Equals("Open") && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                                         && ((item["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                                         && today <= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                         || (item["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                                         && today <= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                         select item)
                                   .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                           where item2["MOC_x0020_Status"].ToString().Equals("Open") && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                           && ((item2["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                           && today <= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                           || (item2["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                           && today <= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                           select item2)
                                   .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                           where item3["MOC_x0020_Status"].ToString().Equals("Open") && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                           && ((item3["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                           && today <= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                           || (item3["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                           && today <= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                           select item3)
                                    .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                            where item4["MOC_x0020_Status"].ToString().Equals("Open") && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && ((item4["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                            || (item4["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                            select item4)
                                    .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                            where item5["MOC_x0020_Status"].ToString().Equals("Open") && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && ((item5["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                            || (item5["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                            select item5)
                                    .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                            where item6["MOC_x0020_Status"].ToString().Equals("Open") && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                             && today >= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddMonths(6).AddDays(-14)
                                             && today <= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddMonths(6)
                                            //&& item6["Current_x0020_Section"].ToString() != "Section 5"
                                            select item6)
                                    .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                            where item7["MOC_x0020_Status"].ToString().Equals("Open") && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && ((item7["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                            || (item7["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                            select item7)
                                    .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                            where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && ((item8["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                            || (item8["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                            select item8)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "Initiate")
                                            && today >= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "First")
                                           && today >= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "First")
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "Second")
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Extension"].ToString() == "Initiate"
                                            && today >= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4).AddDays(-14)
                                            && today <= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Extension"].ToString() == "First"
                                            && today >= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4).AddDays(-14)
                                            && today <= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Extension"].ToString() == "First"
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Extension"].ToString() == "Second"
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                            .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                    where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                    && ((item8["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                    && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                    || (item8["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                    && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                    select item8)
                                   .GroupBy(s => new { Area = s["Area"] })
                                   .OrderBy(s => s.Key.Area)
                                   .Select(s => new
                                   {
                                       areaGroup = s.Key.Area,
                                       Count = s.Count()
                                   });
                        Logger("Area A3 OD 2W");
                        #endregion

                        #region query for Going to NC 2 weeks
                        var getGoingtoNC2weeks = (from item in dt04.Rows.Cast<DataRow>()
                                                  where item["MOC_x0020_Status"].ToString().Equals("Open") && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                                  && today >= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                                  && today <= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"]))
                                                  && item["Current_x0020_Section"].ToString() != "Section 5"
                                                  select item)
                                  .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                          where item2["MOC_x0020_Status"].ToString().Equals("Open") && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                           && today >= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                           && today <= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"]))
                                           && item2["Current_x0020_Section"].ToString() != "Section 5"
                                          select item2)
                                  .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                          where item3["MOC_x0020_Status"].ToString().Equals("Open") && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                           && today >= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                           && today <= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"]))
                                           && item3["Current_x0020_Section"].ToString() != "Section 5"
                                          select item3)
                                   .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                           where item4["MOC_x0020_Status"].ToString().Equals("Open") && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today >= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"]))
                                            && item4["Current_x0020_Section"].ToString() != "Section 5"
                                           select item4)
                                   .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                           where item5["MOC_x0020_Status"].ToString().Equals("Open") && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today >= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"]))
                                            && item5["Current_x0020_Section"].ToString() != "Section 5"
                                           select item5)
                                   .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                           where item6["MOC_x0020_Status"].ToString().Equals("Open") && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today >= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddDays(90).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddDays(90)
                                            && item6["Current_x0020_Section"].ToString() != "Section 5"
                                           select item6)
                                   .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                           where item7["MOC_x0020_Status"].ToString().Equals("Open") && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today >= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"]))
                                            && item7["Current_x0020_Section"].ToString() != "Section 5"
                                           select item7)
                                   .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                           where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"]))
                                            && item8["Current_x0020_Section"].ToString() != "Section 5"
                                           select item8)
                                  .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                          where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                          && item9["Extension"].ToString() == ""
                                          && (item9["Current_x0020_Section"].ToString().Equals("Section 1")
                                          || item9["Current_x0020_Section"].ToString().Equals("Section 2")
                                          || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                          && today >= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])
                                          && today <= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(14)
                                          select item9)
                            .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                    where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                     && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                     && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"]))
                                     && item8["Current_x0020_Section"].ToString() != "Section 5"
                                    select item8)
                                  .GroupBy(s => new { Area = s["Area"] })
                                  .OrderBy(s => s.Key.Area)
                                  .Select(s => new
                                  {
                                      areaGroup = s.Key.Area,
                                      Count = s.Count()
                                  });
                        Logger("Area A3 NC 2W");
                        #endregion

                        #region query for Going to Overdue 1 mth
                        var getGoingToOverdue1mth = (from item in dt04.Rows.Cast<DataRow>()
                                                     where item["MOC_x0020_Status"].ToString().Equals("Open") && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                                     && ((item["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                                     && today <= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                     || (item["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                                     && today <= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                     select item)
                                    .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                            where item2["MOC_x0020_Status"].ToString().Equals("Open") && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && ((item2["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                            && today <= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                            || (item2["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                            && today <= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                            select item2)
                                    .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                            where item3["MOC_x0020_Status"].ToString().Equals("Open") && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && ((item3["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                            && today <= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                            || (item3["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                            && today <= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                            select item3)
                                    .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                            where item4["MOC_x0020_Status"].ToString().Equals("Open") && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && ((item4["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                            && today <= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                            || (item4["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                            && today <= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                            select item4)
                                    .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                            where item5["MOC_x0020_Status"].ToString().Equals("Open") && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && ((item5["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                            && today <= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                            || (item5["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                            && today <= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                            select item5)
                                    .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                            where item6["MOC_x0020_Status"].ToString().Equals("Open") && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                             && today >= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddMonths(5)
                                             && today <= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddMonths(6)
                                            //&& item6["Current_x0020_Section"].ToString() != "Section 5"
                                            select item6)
                                    .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                            where item7["MOC_x0020_Status"].ToString().Equals("Open") && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && ((item7["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                            && today <= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                            || (item7["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                            && today <= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                            select item7)
                                    .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                            where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && ((item8["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                            && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                            || (item8["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                            && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                            select item8)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "Initiate")
                                            && today >= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                            && today <= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "First")
                                           && today >= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                            && today <= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "First")
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "Second")
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Extension"].ToString() == "Initiate"
                                            && today >= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(3)
                                            && today <= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Extension"].ToString() == "First"
                                            && today >= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(3)
                                            && today <= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Extension"].ToString() == "First"
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Extension"].ToString() == "Second"
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                            .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                    where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                    && ((item8["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                    && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                    || (item8["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                    && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                    select item8)
                                   .GroupBy(s => new { Area = s["Area"] })
                                   .OrderBy(s => s.Key.Area)
                                   .Select(s => new
                                   {
                                       areaGroup = s.Key.Area,
                                       Count = s.Count()
                                   });
                        Logger("Area A3 OD 1M");
                        #endregion

                        #region query for Going to NC 1 mth
                        var getGointToNC1mth = (from item in dt04.Rows.Cast<DataRow>()
                                                where item["MOC_x0020_Status"].ToString().Equals("Open") && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && today >= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                                && today <= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"]))
                                                && item["Current_x0020_Section"].ToString() != "Section 5"
                                                select item)
                                   .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                           where item2["MOC_x0020_Status"].ToString().Equals("Open") && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today >= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"]))
                                            && item2["Current_x0020_Section"].ToString() != "Section 5"
                                           select item2)
                                   .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                           where item3["MOC_x0020_Status"].ToString().Equals("Open") && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today >= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"]))
                                            && item3["Current_x0020_Section"].ToString() != "Section 5"
                                           select item3)
                                   .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                           where item4["MOC_x0020_Status"].ToString().Equals("Open") && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today >= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"]))
                                            && item4["Current_x0020_Section"].ToString() != "Section 5"
                                           select item4)
                                   .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                           where item5["MOC_x0020_Status"].ToString().Equals("Open") && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today >= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"]))
                                            && item5["Current_x0020_Section"].ToString() != "Section 5"
                                           select item5)
                                   .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                           where item6["MOC_x0020_Status"].ToString().Equals("Open") && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today >= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddDays(90).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddDays(90)
                                            && item6["Current_x0020_Section"].ToString() != "Section 5"
                                           select item6)
                                   .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                           where item7["MOC_x0020_Status"].ToString().Equals("Open") && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today >= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"]))
                                            && item7["Current_x0020_Section"].ToString() != "Section 5"
                                           select item7)
                                   .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                           where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"]))
                                            && item8["Current_x0020_Section"].ToString() != "Section 5"
                                           select item8)
                                   .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                           where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                           && item9["Extension"].ToString() == ""
                                           && (item9["Current_x0020_Section"].ToString().Equals("Section 1")
                                           || item9["Current_x0020_Section"].ToString().Equals("Section 2")
                                           || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                           && today >= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddDays(14).AddMonths(-1)
                                           && today <= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(14)
                                           select item9)
                            .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                    where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                     && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                     && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"]))
                                     && item8["Current_x0020_Section"].ToString() != "Section 5"
                                    select item8)
                                  .GroupBy(s => new { Area = s["Area"] })
                                  .OrderBy(s => s.Key.Area)
                                  .Select(s => new
                                  {
                                      areaGroup = s.Key.Area,
                                      Count = s.Count()
                                  });
                        Logger("Area A3 NC 1M");
                        #endregion

                        #region query for Overdue Extension OLS
                        var getExtOLS = (from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") 
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "Initiate" || item9["Extension"].ToString() == "")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today > (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(6)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") 
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "First")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today > (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(12)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") 
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "Second")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today > (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(18)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && (item9["Extension"].ToString() == "Initiate" || item9["Extension"].ToString() == "")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today > (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && (item9["Extension"].ToString() == "First")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today > (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(5)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && (item9["Extension"].ToString() == "Second")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today > (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(6)
                                            select item9)
                                   .GroupBy(s => new { Area = s["Area"] })
                                   .OrderBy(s => s.Key.Area)
                                   .Select(s => new
                                   {
                                       areaGroup = s.Key.Area,
                                       Count = s.Count()
                                   });
                        Logger("Area A3 Ext OLS");
                        #endregion

                        #region query for Going to 1 Month Overdue Extension OLS
                        var getExtOLS1M = (from item9 in dt09.Rows.Cast<DataRow>()
                                         where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                         && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                         && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                         && (item9["Extension"].ToString() == "Initiate" || item9["Extension"].ToString() == "")
                                         && !item9["Actual_x0020_Date"].ToString().Equals("")
                                         && today >= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(5)
                                         && today <= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(6)
                                         select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "First")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today >= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(11)
                                            && today <= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(12)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "Second")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today >= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(17)
                                            && today <= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(18)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && (item9["Extension"].ToString() == "Initiate" || item9["Extension"].ToString() == "")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today >= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(4).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && (item9["Extension"].ToString() == "First")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today >= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(5).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(5)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && (item9["Extension"].ToString() == "Second")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today >= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(6).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(6)
                                            select item9)
                                   .GroupBy(s => new { Area = s["Area"] })
                                   .OrderBy(s => s.Key.Area)
                                   .Select(s => new
                                   {
                                       areaGroup = s.Key.Area,
                                       Count = s.Count()
                                   });
                        Logger("Area A3 Ext OLS 1M");
                        #endregion

                        tableEmocCAF = new DataTable();
                        tableEmocCAF.Columns.Add("Area");
                        tableEmocCAF.Columns.Add("Open");
                        tableEmocCAF.Columns.Add("Open2");
                        tableEmocCAF.Columns.Add("Close");
                        tableEmocCAF.Columns.Add("Overdue");
                        tableEmocCAF.Columns.Add("Non_Compliance");
                        tableEmocCAF.Columns.Add("G_T_Overdue_2Weeks");
                        tableEmocCAF.Columns.Add("G_T_NC_2Weeks");
                        tableEmocCAF.Columns.Add("G_T_Overdue_1mth");
                        tableEmocCAF.Columns.Add("G_T_NC_1mth");
                        tableEmocCAF.Columns.Add("Ext_OLS");
                        tableEmocCAF.Columns.Add("Ext_OLS_1m");

                        String[] areaList = { "Area 1- Sweet Hydroskimming", "Area 2- Sour Hydroskimming", "Area 3- Sour Conversion", "Area 4- Storage & Distribution", "Area 5- Utilities & Sulphur Complex", "Area 6- Melaka Group 3 Lube Base Oil", "Non Process" };

                        #region list rearrange open caf type1
                        List<getOpenAreaCAF> open = new List<getOpenAreaCAF>();
                        foreach (var item in getOpenCAF)
                        {

                            getOpenAreaCAF getCAF = new getOpenAreaCAF();
                            getCAF.area = item.areaGroup.ToString();
                            getCAF.count = item.Count.ToString();
                            open.Add(getCAF);
                        }

                        if (open.Count() < areaList.Length)
                        {
                            for (int i = open.Count(); i < areaList.Length; i++)
                            {
                                getOpenAreaCAF getCAF = new getOpenAreaCAF();
                                getCAF.area = "";
                                getCAF.count = "0";
                                open.Add(getCAF);
                            }
                        }
                        #endregion

                        #region rearrange open caf type2
                        List<getOpenAreaCAF> open2 = new List<getOpenAreaCAF>();
                        foreach (var item in getOpenCAF2)
                        {

                            getOpenAreaCAF getCAF = new getOpenAreaCAF();
                            getCAF.area = item.areaGroup.ToString();
                            getCAF.count = item.Count.ToString();
                            open2.Add(getCAF);
                        }

                        if (open2.Count() < areaList.Length)
                        {
                            for (int i = open2.Count(); i < areaList.Length; i++)
                            {
                                getOpenAreaCAF getCAF = new getOpenAreaCAF();
                                getCAF.area = "";
                                getCAF.count = "0";
                                open2.Add(getCAF);
                            }
                        }
                        #endregion

                        #region list rearrange close caf
                        List<getCloseAreaCAF> close = new List<getCloseAreaCAF>();
                        var totalQuery = getCloseCAF.ToArray();
                        foreach (var item in getCloseCAF)
                        {

                            getCloseAreaCAF getCAF = new getCloseAreaCAF();
                            getCAF.area = item.areaGroup.ToString();
                            getCAF.count = item.Count.ToString();
                            close.Add(getCAF);
                        }

                        if (close.Count() < areaList.Length)
                        {
                            for (int i = close.Count(); i < areaList.Length; i++)
                            {
                                getCloseAreaCAF getCAF = new getCloseAreaCAF();
                                getCAF.area = "";
                                getCAF.count = "0";
                                close.Add(getCAF);
                            }
                        }
                        #endregion

                        #region list rearrange overdue caf
                        List<getOverdueAreaCAF> overdue = new List<getOverdueAreaCAF>();
                        foreach (var item in getOverdueCAF)
                        {

                            getOverdueAreaCAF getCAF = new getOverdueAreaCAF();
                            getCAF.area = item.areaGroup.ToString();
                            getCAF.count = item.Count.ToString();
                            overdue.Add(getCAF);
                        }

                        if (overdue.Count() < areaList.Length)
                        {
                            for (int i = overdue.Count(); i < areaList.Length; i++)
                            {
                                getOverdueAreaCAF getCAF = new getOverdueAreaCAF();
                                getCAF.area = "";
                                getCAF.count = "0";
                                overdue.Add(getCAF);
                            }
                        }
                        Logger("Area A4 OD");
                        #endregion

                        #region list rearrange nc caf
                        List<getNCAreaCAF> nc = new List<getNCAreaCAF>();
                        foreach (var item in getNCCAF)
                        {

                            getNCAreaCAF getCAF = new getNCAreaCAF();
                            getCAF.area = item.areaGroup.ToString();
                            getCAF.count = item.Count.ToString();
                            nc.Add(getCAF);
                        }

                        if (nc.Count() < areaList.Length)
                        {
                            for (int i = nc.Count(); i < areaList.Length; i++)
                            {
                                getNCAreaCAF getCAF = new getNCAreaCAF();
                                getCAF.area = "";
                                getCAF.count = "0";
                                nc.Add(getCAF);
                            }
                        }
                        Logger("Area A4 NC");
                        #endregion

                        #region list rearrange gt overdue 2 weeks
                        List<getGTOverdue2weeks> gtOverdue2wks = new List<getGTOverdue2weeks>();
                        foreach (var item in getGoingToOverdue2weekCAF)
                        {

                            getGTOverdue2weeks getCAF = new getGTOverdue2weeks();
                            getCAF.area = item.areaGroup.ToString();
                            getCAF.count = item.Count.ToString();
                            gtOverdue2wks.Add(getCAF);
                        }

                        if (gtOverdue2wks.Count() < areaList.Length)
                        {
                            for (int i = gtOverdue2wks.Count(); i < areaList.Length; i++)
                            {
                                getGTOverdue2weeks getCAF = new getGTOverdue2weeks();
                                getCAF.area = "";
                                getCAF.count = "0";
                                gtOverdue2wks.Add(getCAF);
                            }
                        }
                        Logger("Area A4 OD 2W");
                        #endregion

                        #region list rearrange gt nc 2 weeks
                        List<getGTNC2weeks> gtNC2wks = new List<getGTNC2weeks>();
                        foreach (var item in getGoingtoNC2weeks)
                        {

                            getGTNC2weeks getCAF = new getGTNC2weeks();
                            getCAF.area = item.areaGroup.ToString();
                            getCAF.count = item.Count.ToString();
                            gtNC2wks.Add(getCAF);
                        }

                        if (gtNC2wks.Count() < areaList.Length)
                        {
                            for (int i = gtNC2wks.Count(); i < areaList.Length; i++)
                            {
                                getGTNC2weeks getCAF = new getGTNC2weeks();
                                getCAF.area = "";
                                getCAF.count = "0";
                                gtNC2wks.Add(getCAF);
                            }
                        }
                        Logger("Area A4 NC 2W");
                        #endregion

                        #region list rearrange gt overdue 1 mth
                        List<getGTOverdue1mth> gtOverdue1mth = new List<getGTOverdue1mth>();
                        foreach (var item in getGoingToOverdue1mth)
                        {

                            getGTOverdue1mth getCAF = new getGTOverdue1mth();
                            getCAF.area = item.areaGroup.ToString();
                            getCAF.count = item.Count.ToString();
                            gtOverdue1mth.Add(getCAF);
                        }

                        if (gtOverdue1mth.Count() < areaList.Length)
                        {
                            for (int i = gtOverdue1mth.Count(); i < areaList.Length; i++)
                            {
                                getGTOverdue1mth getCAF = new getGTOverdue1mth();
                                getCAF.area = "";
                                getCAF.count = "0";
                                gtOverdue1mth.Add(getCAF);
                            }
                        }
                        Logger("Area A4 OD 1M");
                        #endregion

                        #region list rearrange gt nc 1 mth
                        List<getGTNC1mth> gtNC1mth = new List<getGTNC1mth>();
                        foreach (var item in getGointToNC1mth)
                        {

                            getGTNC1mth getCAF = new getGTNC1mth();
                            getCAF.area = item.areaGroup.ToString();
                            getCAF.count = item.Count.ToString();
                            gtNC1mth.Add(getCAF);
                        }

                        if (gtNC1mth.Count() < areaList.Length)
                        {
                            for (int i = gtNC1mth.Count(); i < areaList.Length; i++)
                            {
                                getGTNC1mth getCAF = new getGTNC1mth();
                                getCAF.area = "";
                                getCAF.count = "0";
                                gtNC1mth.Add(getCAF);
                            }
                        }
                        Logger("Area A4 NC 1M");
                        #endregion

                        #region list rearrange gt ext OLS
                        List<getExtOLS> ExtOLS = new List<getExtOLS>();
                        foreach (var item in getExtOLS)
                        {

                            getExtOLS getCAF = new getExtOLS();
                            getCAF.area = item.areaGroup.ToString();
                            getCAF.count = item.Count.ToString();
                            ExtOLS.Add(getCAF);
                        }

                        if (ExtOLS.Count() < areaList.Length)
                        {
                            for (int i = ExtOLS.Count(); i < areaList.Length; i++)
                            {
                                getExtOLS getCAF = new getExtOLS();
                                getCAF.area = "";
                                getCAF.count = "0";
                                ExtOLS.Add(getCAF);
                            }
                        }
                        Logger("Area A4 Ext OLS");
                        #endregion

                        #region list rearrange gt Ext OLS 1m
                        List<getExtOLS1M> ExtOLS1M = new List<getExtOLS1M>();
                        foreach (var item in getExtOLS1M)
                        {

                            getExtOLS1M getCAF = new getExtOLS1M();
                            getCAF.area = item.areaGroup.ToString();
                            getCAF.count = item.Count.ToString();
                            ExtOLS1M.Add(getCAF);
                        }

                        if (ExtOLS1M.Count() < areaList.Length)
                        {
                            for (int i = ExtOLS1M.Count(); i < areaList.Length; i++)
                            {
                                getExtOLS1M getCAF = new getExtOLS1M();
                                getCAF.area = "";
                                getCAF.count = "0";
                                ExtOLS1M.Add(getCAF);
                            }
                        }
                        Logger("Area A4 Ext OLS 1M");
                        #endregion

                        for (int i = 0; i < areaList.Length; i++)
                        {

                            DataRow dr = tableEmocCAF.NewRow();
                            dr["Area"] = areaList[i].ToString();

                            //open
                            if (open[i].area.ToString() == areaList[i].ToString())
                            {

                                dr["Open"] = open[i].count.ToString();
                            }
                            else
                            {
                                dr["Open"] = "0";
                                if (i < 6)
                                {
                                    open.RemoveAt(6);
                                    open.Insert(i + 1, open[i]);
                                }
                            }

                            //open2
                            if (open2[i].area.ToString() == areaList[i].ToString())
                            {

                                dr["Open2"] = open2[i].count.ToString();
                            }
                            else
                            {
                                dr["Open2"] = "0";
                                if (i < 6)
                                {
                                    open2.RemoveAt(6);
                                    open2.Insert(i + 1, open2[i]);
                                }
                            }

                            //close
                            if (close[i].area.ToString() == areaList[i].ToString())
                            {
                                dr["Close"] = close[i].count.ToString();
                            }
                            else
                            {
                                dr["Close"] = "0";
                                if (i < 6)
                                {
                                    close.RemoveAt(6);
                                    close.Insert(i + 1, close[i]);
                                }

                            }

                            //overdue
                            if (overdue[i].area.ToString() == areaList[i].ToString())
                            {
                                dr["Overdue"] = overdue[i].count.ToString();
                            }
                            else
                            {
                                dr["Overdue"] = "0";
                                if (i < 6)
                                {
                                    overdue.RemoveAt(6);
                                    overdue.Insert(i + 1, overdue[i]);
                                }
                            }

                            //NC
                            if (nc[i].area.ToString() == areaList[i].ToString())
                            {
                                dr["Non_Compliance"] = nc[i].count.ToString();
                            }
                            else
                            {
                                dr["Non_Compliance"] = "0";
                                if (i < 6)
                                {
                                    nc.RemoveAt(6);
                                    nc.Insert(i + 1, nc[i]);
                                }

                            }

                            //GT Overdue 2 weeks
                            if (gtOverdue2wks[i].area.ToString() == areaList[i].ToString())
                            {
                                dr["G_T_Overdue_2Weeks"] = gtOverdue2wks[i].count.ToString();
                            }
                            else
                            {
                                dr["G_T_Overdue_2Weeks"] = "0";
                                if (i < 6)
                                {
                                    gtOverdue2wks.RemoveAt(6);
                                    gtOverdue2wks.Insert(i + 1, gtOverdue2wks[i]);
                                }
                            }

                            //GT NC 2 weeks
                            if (gtNC2wks[i].area.ToString() == areaList[i].ToString())
                            {
                                dr["G_T_NC_2Weeks"] = gtNC2wks[i].count.ToString();
                            }
                            else
                            {
                                dr["G_T_NC_2Weeks"] = "0";
                                if (i < 6)
                                {
                                    gtNC2wks.RemoveAt(6);
                                    gtNC2wks.Insert(i + 1, gtNC2wks[i]);
                                }
                            }

                            //GT Overdue 1 mth
                            if (gtOverdue1mth[i].area.ToString() == areaList[i].ToString())
                            {
                                dr["G_T_Overdue_1mth"] = gtOverdue1mth[i].count.ToString();
                            }
                            else
                            {
                                dr["G_T_Overdue_1mth"] = "0";
                                if (i < 6)
                                {
                                    gtOverdue1mth.RemoveAt(6);
                                    gtOverdue1mth.Insert(i + 1, gtOverdue1mth[i]);
                                }
                            }

                            //GT NC 1 mth
                            if (gtNC1mth[i].area.ToString() == areaList[i].ToString())
                            {
                                dr["G_T_NC_1mth"] = gtNC1mth[i].count.ToString();
                            }
                            else
                            {
                                dr["G_T_NC_1mth"] = "0";
                                if (i < 6)
                                {
                                    gtNC1mth.RemoveAt(6);
                                    gtNC1mth.Insert(i + 1, gtNC1mth[i]);
                                }
                            }

                            //GT Ext OLS
                            if (ExtOLS[i].area.ToString() == areaList[i].ToString())
                            {
                                dr["Ext_OLS"] = ExtOLS[i].count.ToString();
                            }
                            else
                            {
                                dr["Ext_OLS"] = "0";
                                if (i < 6)
                                {
                                    ExtOLS.RemoveAt(6);
                                    ExtOLS.Insert(i + 1, ExtOLS[i]);
                                }
                            }
                            Logger("Area A4 Ext_OLS");

                            //GT Ext OLS 1m
                            if (ExtOLS1M[i].area.ToString() == areaList[i].ToString())
                            {
                                dr["Ext_OLS_1m"] = ExtOLS[i].count.ToString();
                            }
                            else
                            {
                                dr["Ext_OLS_1m"] = "0";
                                if (i < 6)
                                {
                                    ExtOLS1M.RemoveAt(6);
                                    ExtOLS1M.Insert(i + 1, ExtOLS1M[i]);
                                }
                            }
                            Logger("Area A4 Ext_OLS_1M");

                            tableEmocCAF.Rows.Add(dr);
                        }

                        ListViewCAF.DataSource = tableEmocCAF;
                        ListViewCAF.DataBind();
                        Logger("Area A5");
                    }
                    catch (Exception ex)
                    {
                        Logger("Error catch: " + ex.ToString());
                    }
                } 

            }
            //oWeb.Dispose();
            //oSite.Dispose();

        }

        private object Concat(IEnumerable<DataRow> enumerable)
        {
            throw new NotImplementedException();
        }

        protected void ListViewCAF_ItemCommand(object sender, ListViewCommandEventArgs e)
        {

            LinkButton LinkBtn = (LinkButton)e.Item.FindControl("linkOpen");
            LinkButton LinkBtn2 = (LinkButton)e.Item.FindControl("linkOpen2");
            LinkButton LinkClose = (LinkButton)e.Item.FindControl("linkClose");
            LinkButton LinkOverdue = (LinkButton)e.Item.FindControl("linkOverdue");
            LinkButton LinkNC = (LinkButton)e.Item.FindControl("linkNC");
            LinkButton LinkGTO = (LinkButton)e.Item.FindControl("linkGTOverdue2Weeks");
            LinkButton LinkGTNC = (LinkButton)e.Item.FindControl("linkGTNC2weeks");
            LinkButton LinkGTO_1month = (LinkButton)e.Item.FindControl("linkGTOverdue1mth");
            LinkButton LinkGTNC_1month = (LinkButton)e.Item.FindControl("linknkGTNC1mth");
            LinkButton LinkExtOLS = (LinkButton)e.Item.FindControl("linkExt_OLS");
            LinkButton LinkExtOLS_1month = (LinkButton)e.Item.FindControl("linkExt_OLS_1m");
            ListViewCAF.Visible = false;
            ListViewCAFDetails.Visible = true;
            btnBack.Visible = true;

            int area = e.Item.DataItemIndex + 1;

            if (e.CommandName == "Open")// CommandName of LinkButton
            {
                //test.Text = "Area:" + e.Item.DataItemIndex.ToString() + ", Open:" + LinkBtn.Text;
                ShowDetails("Open", area);
            }
            else if (e.CommandName == "Open2")
            {
                ShowDetails("Open2", area);
            }
            else if (e.CommandName == "Closed")
            {
                ShowDetails("Closed", area);
            }
            else if (e.CommandName == "Overdue")
            {
                ShowDetails("Overdue", area);
            }
            else if (e.CommandName == "NC")
            {
                ShowDetails("NC", area);
            }
            else if (e.CommandName == "GTO")
            {
                ShowDetails("GTO", area);
            }
            else if (e.CommandName == "GTNC")
            {
                ShowDetails("GTNC", area);
            }
            else if (e.CommandName == "GTO_1month")
            {
                ShowDetails("GTO_1month", area);
            }
            else if (e.CommandName == "GTNC_1month")
            {
                ShowDetails("GTNC_1month", area);
            }
            else if (e.CommandName == "Ext_OLS")
            {
                ShowDetails("Ext_OLS", area);
            }
            else if (e.CommandName == "Ext_OLS_1m")
            {
                ShowDetails("Ext_OLS_1m", area);
            }
        }

        protected void ShowDetails(string status, int area)
        {
            /*Area 1- Sweet Hydroskimming
            Area 2- Sour Hydroskimming
            Area 3- Sour Conversion
            Area 4- Storage & Distribution
            Area 5- Utilities & Sulphur Complex
            Area 6- Melaka Group 3 Lube Base Oil*/

            string currentArea = "";

            #region Define area
            if (area == 1)
                currentArea = "Area 1- Sweet Hydroskimming";
            else if (area == 2)
                currentArea = "Area 2- Sour Hydroskimming";
            else if (area == 3)
                currentArea = "Area 3- Sour Conversion";
            else if (area == 4)
                currentArea = "Area 4- Storage & Distribution";
            else if (area == 5)
                currentArea = "Area 5- Utilities & Sulphur Complex";
            else if (area == 6)
                currentArea = "Area 6- Melaka Group 3 Lube Base Oil";
            else if (area == 7)
                currentArea = "Non Process";
            #endregion

            string URLwebmoc = "";
            URLwebmoc = SPContext.Current.Web.Url;

            using (SPSite oSite = new SPSite(URLwebmoc))
            {
                using (SPWeb oWeb = oSite.OpenWeb())
                {

                    SPList oList04 = oWeb.Lists["Standard Form (CAF-SUP-04)"];
                    SPList oList04A = oWeb.Lists["Non Process Area Change (CAF-SUP-04A)"];
                    SPList oList04B = oWeb.Lists["Transmitter Re-Ranging (CAF-SUP-04B)"];
                    SPList oList04C = oWeb.Lists["DCS Minor Change (CAF-SUP-04C)"];
                    SPList oList04D = oWeb.Lists["DCS Non Critical Alarm Parameter Change (CAF-SUP-04D)"];
                    SPList oList04E = oWeb.Lists["SCE Bypass (CAF-SUP-04E)"];
                    SPList oList04G = oWeb.Lists["Product Change (CAF-SUP-04G)"];
                    SPList oList04J = oWeb.Lists["Process Control Network Change (for Level 3 and above) (CAF-SUP-04J)"];
                    SPList oList09 = oWeb.Lists["Online Leak Sealing (CAF-SUP-09)"];
                    SPList oList04L = oWeb.Lists["Rejuvenate Change (CAF-SUP-04L)"];
                    //CAFSUP04L to be added
                    //SPList oList04L= oWeb.Lists["Total Status Report"];

                    SPListItemCollection item04 = oList04.GetItems("CAF_x0020_No", "CRP_x0020_Name", "Title", "Construction_Date", "Date_x0020_Initiated", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF"); //Items //nacb_09042020_override_remarks
                    SPListItemCollection item04A = oList04A.GetItems("CAF_x0020_No", "CRP_x0020_Name", "Title", "Construction_Date", "Date_x0020_Initiated", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF"); //.GetItems("CAF_x0020_No", "CAF_x0020_Name", "Title", "Date_x0020_Initiated", "Est_x002e__x0020_Start_x0020_Up_", "Modified", "MOC_x0020_Status", "Area");
                    SPListItemCollection item04B = oList04B.GetItems("CAF_x0020_No", "CRP_x0020_Name", "Title", "Construction_Date", "Date_x0020_Initiated", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF");
                    SPListItemCollection item04C = oList04C.GetItems("CAF_x0020_No", "CRP_x0020_Name", "Title", "Construction_Date", "Date_x0020_Initiated", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF");
                    SPListItemCollection item04D = oList04D.GetItems("CAF_x0020_No", "CRP_x0020_Name", "Title", "Construction_Date", "Date_x0020_Initiated", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF");
                    SPListItemCollection item04E = oList04E.GetItems("CAF_x0020_No", "CRP_x0020_Name", "Title", "Construction_Date", "Date_x0020_Initiated", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "ByPass_x0020_Date");
                    SPListItemCollection item04G = oList04G.GetItems("CAF_x0020_No", "CRP_x0020_Name", "Title", "Construction_Date", "Date_x0020_Initiated", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF");
                    SPListItemCollection item04J = oList04J.GetItems("CAF_x0020_No", "CRP_x0020_Name", "Title", "Construction_Date", "Date_x0020_Initiated", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF");
                    SPListItemCollection item09 = oList09.GetItems("CAF_x0020_No", "CRP_x0020_Name", "Title", "Construction_Date", "Date_x0020_Initiated", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Type_x0020_of_x0020_Repair", "Extension", "Extension_Date", "Override_x0020_Remarks", "Actual_x0020_Date", "Sect3_HOP_App", "Sect4A_x0020_AO_x0020_Approval_x");
                    SPListItemCollection item04L = oList04L.GetItems("CAF_x0020_No", "CRP_x0020_Name", "Title", "Construction_Date", "Date_x0020_Initiated", "MOC_x0020_Status", "Area", "DATA_x0020_TYPE", "Est_x002e__x0020_Start_x0020_Up_", "Current_x0020_Section", "Override_x0020_Remarks", "Affected_CAF");
                    Logger("Area A6");
                    //DataTable dt04;//not initialized, will be null

                    DataTable dt04;//initialized
                    DataTable dt04A;
                    DataTable dt04B;
                    DataTable dt04C;
                    DataTable dt04D;
                    DataTable dt04E;
                    DataTable dt04G;
                    DataTable dt04J;
                    DataTable dt09;
                    DataTable dt04L;

                    #region Declare DataTable

                    if (item04.Count > 0)
                    {
                        dt04 = item04.GetDataTable();
                        dt04.Columns.Add("CAFType", typeof(String));
                    }
                    else
                    {
                        dt04 = new DataTable();

                        dt04.Columns.Add("CAF_x0020_No");
                        dt04.Columns.Add("CRP_x0020_Name");
                        dt04.Columns.Add("Title");
                        dt04.Columns.Add("Construction_Date");
                        dt04.Columns.Add("Date_x0020_Initiated");
                        dt04.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04.Columns.Add("CAFType");
                        dt04.Columns.Add("DATA_x0020_TYPE");
                        dt04.Columns.Add("Override_x0020_Remarks");
                        dt04.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A7 04");

                    if (item04A.Count > 0)
                    {
                        dt04A = item04A.GetDataTable();
                        dt04A.Columns.Add("CAFType", typeof(String));
                    }
                    else
                    {
                        dt04A = new DataTable();

                        dt04A.Columns.Add("CAF_x0020_No");
                        dt04A.Columns.Add("CRP_x0020_Name");
                        dt04A.Columns.Add("Title");
                        dt04A.Columns.Add("Construction_Date");
                        dt04A.Columns.Add("Date_x0020_Initiated");
                        dt04A.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04A.Columns.Add("CAFType");
                        dt04A.Columns.Add("DATA_x0020_TYPE");
                        dt04A.Columns.Add("Override_x0020_Remarks");
                        dt04A.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A7 04A");


                    if (item04B.Count > 0)
                    {
                        dt04B = item04B.GetDataTable();
                        dt04B.Columns.Add("CAFType", typeof(String));
                    }
                    else
                    {
                        dt04B = new DataTable();

                        dt04B.Columns.Add("CAF_x0020_No");
                        dt04B.Columns.Add("CRP_x0020_Name");
                        dt04B.Columns.Add("Title");
                        dt04B.Columns.Add("Construction_Date");
                        dt04B.Columns.Add("Date_x0020_Initiated");
                        dt04B.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04B.Columns.Add("CAFType");
                        dt04B.Columns.Add("DATA_x0020_TYPE");
                        dt04B.Columns.Add("Override_x0020_Remarks");
                        dt04B.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A7 04B");


                    if (item04C.Count > 0)
                    {
                        dt04C = item04C.GetDataTable();
                        dt04C.Columns.Add("CAFType", typeof(String));
                    }
                    else
                    {
                        dt04C = new DataTable();

                        dt04C.Columns.Add("CAF_x0020_No");
                        dt04C.Columns.Add("CRP_x0020_Name");
                        dt04C.Columns.Add("Title");
                        dt04C.Columns.Add("Construction_Date");
                        dt04C.Columns.Add("Date_x0020_Initiated");
                        dt04C.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04C.Columns.Add("CAFType");
                        dt04C.Columns.Add("DATA_x0020_TYPE");
                        dt04C.Columns.Add("Override_x0020_Remarks");
                        dt04C.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A7 04C");

                    if (item04D.Count > 0)
                    {
                        dt04D = item04D.GetDataTable();
                        dt04D.Columns.Add("CAFType", typeof(String));
                    }
                    else
                    {
                        dt04D = new DataTable();

                        dt04D.Columns.Add("CAF_x0020_No");
                        dt04D.Columns.Add("CRP_x0020_Name");
                        dt04D.Columns.Add("Title");
                        dt04D.Columns.Add("Construction_Date");
                        dt04D.Columns.Add("Date_x0020_Initiated");
                        dt04D.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04D.Columns.Add("CAFType");
                        dt04D.Columns.Add("DATA_x0020_TYPE");
                        dt04D.Columns.Add("Override_x0020_Remarks");
                        dt04D.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A8 04D");

                    if (item04E.Count > 0)
                    {
                        dt04E = item04E.GetDataTable();
                        dt04E.Columns.Add("CAFType", typeof(String));
                    }
                    else
                    {
                        dt04E = new DataTable();

                        dt04E.Columns.Add("CAF_x0020_No");
                        dt04E.Columns.Add("CRP_x0020_Name");
                        dt04E.Columns.Add("Title");
                        dt04E.Columns.Add("Construction_Date");
                        dt04E.Columns.Add("Date_x0020_Initiated");
                        dt04E.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04E.Columns.Add("CAFType");
                        dt04E.Columns.Add("DATA_x0020_TYPE");
                        dt04E.Columns.Add("Override_x0020_Remarks");
                        dt04E.Columns.Add("ByPass_x0020_Date");
                    }
                    Logger("Area A9 04E");

                    if (item04G.Count > 0)
                    {
                        dt04G = item04G.GetDataTable();
                        dt04G.Columns.Add("CAFType", typeof(String));
                    }
                    else
                    {
                        dt04G = new DataTable();

                        dt04G.Columns.Add("CAF_x0020_No");
                        dt04G.Columns.Add("CRP_x0020_Name");
                        dt04G.Columns.Add("Title");
                        dt04G.Columns.Add("Construction_Date");
                        dt04G.Columns.Add("Date_x0020_Initiated");
                        dt04G.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04G.Columns.Add("CAFType");
                        dt04G.Columns.Add("DATA_x0020_TYPE");
                        dt04G.Columns.Add("Override_x0020_Remarks");
                        dt04G.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A9 04G");


                    if (item04J.Count > 0)
                    {
                        dt04J = item04J.GetDataTable();
                        dt04J.Columns.Add("CAFType", typeof(String));
                    }
                    else
                    {
                        dt04J = new DataTable();

                        dt04J.Columns.Add("CAF_x0020_No");
                        dt04J.Columns.Add("CRP_x0020_Name");
                        dt04J.Columns.Add("Title");
                        dt04J.Columns.Add("Construction_Date");
                        dt04J.Columns.Add("Date_x0020_Initiated");
                        dt04J.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04J.Columns.Add("CAFType");
                        dt04J.Columns.Add("DATA_x0020_TYPE");
                        dt04J.Columns.Add("Override_x0020_Remarks");
                        dt04J.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A9 04J");

                    if (item09.Count > 0)
                    {
                        dt09 = item09.GetDataTable();
                        dt09.Columns.Add("CAFType", typeof(String));
                    }
                    else
                    {
                        dt09 = new DataTable();

                        dt09.Columns.Add("CAF_x0020_No");
                        dt09.Columns.Add("CRP_x0020_Name");
                        dt09.Columns.Add("Title");
                        dt09.Columns.Add("Construction_Date");
                        dt09.Columns.Add("Date_x0020_Initiated");
                        dt09.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt09.Columns.Add("CAFType");
                        dt09.Columns.Add("DATA_x0020_TYPE");
                        dt09.Columns.Add("Type_x0020_of_x0020_Repair");
                        dt09.Columns.Add("Extension");
                        dt09.Columns.Add("Extension_Date");
                        dt09.Columns.Add("Override_x0020_Remarks");
                        dt09.Columns.Add("Actual_x0020_Date");
                        dt09.Columns.Add("Sect3_HOP_App");
                        dt09.Columns.Add("Sect4A_x0020_AO_x0020_Approval_x");
                    }
                    Logger("Area A9 09");

                    if (item04L.Count > 0)
                    {
                        dt04L = item04L.GetDataTable();
                        dt04L.Columns.Add("CAFType", typeof(String));
                    }
                    else
                    {
                        dt04L = new DataTable();

                        dt04L.Columns.Add("CAF_x0020_No");
                        dt04L.Columns.Add("CRP_x0020_Name");
                        dt04L.Columns.Add("Title");
                        dt04L.Columns.Add("Construction_Date");
                        dt04L.Columns.Add("Date_x0020_Initiated");
                        dt04L.Columns.Add("Est_x002e__x0020_Start_x0020_Up_");
                        dt04L.Columns.Add("CAFType");
                        dt04L.Columns.Add("DATA_x0020_TYPE");
                        dt04L.Columns.Add("Override_x0020_Remarks");
                        dt04L.Columns.Add("Affected_CAF");
                    }
                    Logger("Area A9 04L");

                    #endregion

                    #region assign CAF TYPE

                    /**/
                    foreach (DataRow row in dt04.Rows)
                    {
                        row["CAFType"] = "Standard Form (CAF-SUP-04)";  // or set it to some other value
                    }
                    dt04.AcceptChanges();

                    foreach (DataRow row in dt04A.Rows)
                    {
                        row["CAFType"] = "Non Process Area Change (CAF-SUP-04A)";  // or set it to some other value
                    }
                    dt04A.AcceptChanges();

                    foreach (DataRow row in dt04B.Rows)
                    {
                        row["CAFType"] = "Transmitter Re-Ranging (CAF-SUP-04B)";  // or set it to some other value
                    }
                    dt04B.AcceptChanges();

                    foreach (DataRow row in dt04C.Rows)
                    {
                        row["CAFType"] = "DCS Minor Change (CAF-SUP-04C)";  // or set it to some other value
                    }
                    dt04C.AcceptChanges();

                    foreach (DataRow row in dt04D.Rows)
                    {
                        row["CAFType"] = "DCS Non Critical Alarm Parameter Change (CAF-SUP-04D)";  // or set it to some other value
                    }
                    dt04D.AcceptChanges();

                    foreach (DataRow row in dt04E.Rows)
                    {
                        row["CAFType"] = "SCE Bypass (CAF-SUP-04E)";  // or set it to some other value
                    }
                    dt04E.AcceptChanges();

                    foreach (DataRow row in dt04G.Rows)
                    {
                        row["CAFType"] = "Product Change (CAF-SUP-04G)";  // or set it to some other value
                    }
                    dt04G.AcceptChanges();

                    foreach (DataRow row in dt04J.Rows)
                    {
                        row["CAFType"] = "Process Control Network Change (for Level 3 and above) (CAF-SUP-04J)";  // or set it to some other value
                    }
                    dt04J.AcceptChanges();

                    foreach (DataRow row in dt09.Rows)
                    {
                        row["CAFType"] = "Online Leak Sealing (CAF-SUP-09)";  // or set it to some other value
                    }
                    dt09.AcceptChanges();

                    foreach (DataRow row in dt04L.Rows)
                    {
                        row["CAFType"] = "Rejuvenate Change (CAF-SUP-04L)";  // or set it to some other value
                    }
                    dt04L.AcceptChanges();

                    #endregion
                    //DataTable tableEmocCAF;
                    // DataView tableEmocCAFView;

                    var today = DateTime.Today;
                    //var getCAF;

                    try
                    {

                        if (status.Equals("Open"))
                        {
                            #region Query Open CAF
                            //query for Open CAF
                            var getOpenCAF = (from item in dt04.Rows.Cast<DataRow>() //item04.OfType<SPListItem>()
                                              where item["MOC_x0020_Status"].ToString().Equals("Open")
                                              && item["Area"].ToString().Equals(currentArea)
                                              && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                              && item["Current_x0020_Section"].ToString() != "Section 5"
                                              select item)
                                              .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                                      where item2["MOC_x0020_Status"].ToString().Equals("Open")
                                                      && item2["Area"].ToString().Equals(currentArea)
                                                      && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                                      && item2["Current_x0020_Section"].ToString() != "Section 5"
                                                      select item2)
                                               .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                                       where item3["MOC_x0020_Status"].ToString().Equals("Open")
                                                      && item3["Area"].ToString().Equals(currentArea)
                                                      && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                                      && item3["Current_x0020_Section"].ToString() != "Section 5"
                                                       select item3)
                                               .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                                       where item4["MOC_x0020_Status"].ToString().Equals("Open")
                                                       && item4["Area"].ToString().Equals(currentArea)
                                                       && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                                       && item4["Current_x0020_Section"].ToString() != "Section 5"
                                                       select item4)
                                               .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                                       where item5["MOC_x0020_Status"].ToString().Equals("Open")
                                                       && item5["Area"].ToString().Equals(currentArea)
                                                       && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                                       && item5["Current_x0020_Section"].ToString() != "Section 5"
                                                       select item5)
                                               .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                                       where item6["MOC_x0020_Status"].ToString().Equals("Open")
                                                       && item6["Area"].ToString().Equals(currentArea)
                                                       && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                                       && item6["Current_x0020_Section"].ToString() != "Section 5"
                                                       select item6)
                                               .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                                       where item7["MOC_x0020_Status"].ToString().Equals("Open")
                                                       && item7["Area"].ToString().Equals(currentArea)
                                                       && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                                       && item7["Current_x0020_Section"].ToString() != "Section 5"
                                                       select item7)
                                               .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                                       where item8["MOC_x0020_Status"].ToString().Equals("Open")
                                                       && item8["Area"].ToString().Equals(currentArea)
                                                       && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                                       && item8["Current_x0020_Section"].ToString() != "Section 5"
                                                       select item8)
                                               .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                                       where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                                       && item9["Area"].ToString().Equals(currentArea)
                                                       && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                                       && item9["Current_x0020_Section"].ToString() != "Section 5"
                                                       select item9)
                                .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                        where item8["MOC_x0020_Status"].ToString().Equals("Open")
                                        && item8["Area"].ToString().Equals(currentArea)
                                        && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                        && item8["Current_x0020_Section"].ToString() != "Section 5"
                                        select item8)
                                               .Select(s => new
                                               {
                                                   CafNo = s["CAF_x0020_No"],
                                                   CAFType = s["CAFType"],
                                                   Title = s["Title"],
                                                   CRPName = s["CRP_x0020_Name"],
                                                   CurrentSection = s["Current_x0020_Section"],
                                                   InitiationDate = s["Date_x0020_Initiated"],
                                                   ConstructionDate = s["Construction_Date"],
                                                   StartUpDate = s["Est_x002e__x0020_Start_x0020_Up_"],
                                                   SubmissionDate = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6),
                                                   Remarks = s["Override_x0020_Remarks"]
                                               });
                            ListViewCAFDetails.DataSource = getOpenCAF;
                            ListViewCAFDetails.DataBind();

                            #endregion
                        }

                        if (status.Equals("Open2"))
                        {
                            #region Query Open CAF
                            //query for Open CAF
                            var getOpenCAF = (from item in dt04.Rows.Cast<DataRow>() //item04.OfType<SPListItem>()
                                              where item["MOC_x0020_Status"].ToString().Equals("Open")
                                              && item["Area"].ToString().Equals(currentArea)
                                              && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                              && item["Current_x0020_Section"].ToString() == "Section 5"
                                              select item)
                                               .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                                       where item2["MOC_x0020_Status"].ToString().Equals("Open")
                                                       && item2["Area"].ToString().Equals(currentArea)
                                                       && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                                       && item2["Current_x0020_Section"].ToString() == "Section 5"
                                                       select item2)
                                                .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                                        where item3["MOC_x0020_Status"].ToString().Equals("Open")
                                                       && item3["Area"].ToString().Equals(currentArea)
                                                       && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                                       && item3["Current_x0020_Section"].ToString() == "Section 5"
                                                        select item3)
                                               .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                                       where item4["MOC_x0020_Status"].ToString().Equals("Open")
                                                       && item4["Area"].ToString().Equals(currentArea)
                                                       && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                                       && item4["Current_x0020_Section"].ToString() == "Section 5"
                                                       select item4)
                                               .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                                       where item5["MOC_x0020_Status"].ToString().Equals("Open")
                                                       && item5["Area"].ToString().Equals(currentArea)
                                                       && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                                       && item5["Current_x0020_Section"].ToString() == "Section 5"
                                                       select item5)
                                               .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                                       where item6["MOC_x0020_Status"].ToString().Equals("Open")
                                                       && item6["Area"].ToString().Equals(currentArea)
                                                       && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                                       && item6["Current_x0020_Section"].ToString() == "Section 5"
                                                       select item6)
                                               .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                                       where item7["MOC_x0020_Status"].ToString().Equals("Open")
                                                       && item7["Area"].ToString().Equals(currentArea)
                                                       && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                                       && item7["Current_x0020_Section"].ToString() == "Section 5"
                                                       select item7)
                                               .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                                       where item8["MOC_x0020_Status"].ToString().Equals("Open")
                                                       && item8["Area"].ToString().Equals(currentArea)
                                                       && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                                       && item8["Current_x0020_Section"].ToString() == "Section 5"
                                                       select item8)
                                               .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                                       where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                                       && item9["Area"].ToString().Equals(currentArea)
                                                       && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                                       && item9["Current_x0020_Section"].ToString() == "Section 5"
                                                       select item9)
                                .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                        where item8["MOC_x0020_Status"].ToString().Equals("Open")
                                        && item8["Area"].ToString().Equals(currentArea)
                                        && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                        && item8["Current_x0020_Section"].ToString() == "Section 5"
                                        select item8)
                                               .Select(s => new
                                               {
                                                   CafNo = s["CAF_x0020_No"],
                                                   CAFType = s["CAFType"],
                                                   Title = s["Title"],
                                                   CRPName = s["CRP_x0020_Name"],
                                                   CurrentSection = s["Current_x0020_Section"],
                                                   InitiationDate = s["Date_x0020_Initiated"],
                                                   ConstructionDate = s["Construction_Date"],
                                                   StartUpDate = s["Est_x002e__x0020_Start_x0020_Up_"],
                                                   SubmissionDate = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6),
                                                   Remarks = s["Override_x0020_Remarks"]
                                               });
                            ListViewCAFDetails.DataSource = getOpenCAF;
                            ListViewCAFDetails.DataBind();

                            #endregion
                        }

                        if (status.Equals("Closed"))
                        {
                            #region Query Closed CAF
                            //query for Close CAF
                            var getCloseCAF =
                                       (from item in dt04.Rows.Cast<DataRow>()
                                        where item["MOC_x0020_Status"].ToString().Equals("Closed")
                                            && item["Area"].ToString().Equals(currentArea)
                                            && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                        select item)
                                      .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                              where item2["MOC_x0020_Status"].ToString().Equals("Closed")
                                               && item2["Area"].ToString().Equals(currentArea)
                                               && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                              select item2)
                                       .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                               where item3["MOC_x0020_Status"].ToString().Equals("Closed")
                                                && item3["Area"].ToString().Equals(currentArea)
                                                && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item3)
                                       .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                               where item4["MOC_x0020_Status"].ToString().Equals("Closed")
                                                && item4["Area"].ToString().Equals(currentArea)
                                                && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item4)
                                       .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                               where item5["MOC_x0020_Status"].ToString().Equals("Closed")
                                                && item5["Area"].ToString().Equals(currentArea)
                                                && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item5)
                                       .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                               where item6["MOC_x0020_Status"].ToString().Equals("Closed")
                                                && item6["Area"].ToString().Equals(currentArea)
                                                && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item6)
                                       .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                               where item7["MOC_x0020_Status"].ToString().Equals("Closed")
                                                && item7["Area"].ToString().Equals(currentArea)
                                                && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item7)
                                       .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                               where item8["MOC_x0020_Status"].ToString().Equals("Closed")
                                                && item8["Area"].ToString().Equals(currentArea)
                                                && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item8)
                                       .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                               where item9["MOC_x0020_Status"].ToString().Equals("Closed")
                                                && item9["Area"].ToString().Equals(currentArea)
                                                && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item9)
                                .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                        where item8["MOC_x0020_Status"].ToString().Equals("Closed")
                                         && item8["Area"].ToString().Equals(currentArea)
                                         && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                        select item8)
                                                   .Select(s => new
                                                   {
                                                       CafNo = s["CAF_x0020_No"],
                                                       CAFType = s["CAFType"],
                                                       Title = s["Title"],
                                                       CRPName = s["CRP_x0020_Name"],
                                                       CurrentSection = s["Current_x0020_Section"],
                                                       InitiationDate = s["Date_x0020_Initiated"],
                                                       ConstructionDate = s["Construction_Date"],
                                                       StartUpDate = s["Est_x002e__x0020_Start_x0020_Up_"],
                                                       SubmissionDate = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6),
                                                       Remarks = s["Override_x0020_Remarks"]
                                                   });
                            ListViewCAFDetails.DataSource = getCloseCAF;
                            ListViewCAFDetails.DataBind();
                            #endregion
                        }
                      
                        if (status.Equals("Overdue"))
                        {
                            #region Query Overdue CAF
                            //query for Overdue CAF
                            var getOverdueCAF = (from item in dt04.Rows.Cast<DataRow>()
                                                 where item["MOC_x0020_Status"].ToString().Equals("Open") && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                                 && ((item["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)) ||
                                                 (item["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                 && item["Area"].ToString().Equals(currentArea)
                                                 //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                                 select item)
                                        .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                                where item2["MOC_x0020_Status"].ToString().Equals("Open") && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                                 && ((item2["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)) ||
                                                 (item2["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                 && item2["Area"].ToString().Equals(currentArea)
                                                //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                                select item2)
                                       .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                               where item3["MOC_x0020_Status"].ToString().Equals("Open") && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                                 && ((item3["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)) ||
                                                 (item3["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                 && item3["Area"].ToString().Equals(currentArea)
                                               //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                               select item3)
                                        .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                                where item4["MOC_x0020_Status"].ToString().Equals("Open") && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                                 && ((item4["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)) ||
                                                 (item4["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                 && item4["Area"].ToString().Equals(currentArea)
                                                //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                                select item4)
                                        .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                                where item5["MOC_x0020_Status"].ToString().Equals("Open") && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                                 && ((item5["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)) ||
                                                 (item5["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                 && item5["Area"].ToString().Equals(currentArea)
                                                //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                                select item5)
                                        .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                                where item6["MOC_x0020_Status"].ToString().Equals("Open") && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                                 && today > (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddMonths(6)
                                                 && item6["Area"].ToString().Equals(currentArea)
                                                //&& item6["Current_x0020_Section"].ToString() != "Section 5"
                                                select item6)
                                        .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                                where item7["MOC_x0020_Status"].ToString().Equals("Open") && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                                 && ((item7["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)) ||
                                                 (item7["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                 && item7["Area"].ToString().Equals(currentArea)
                                                //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                                select item7)
                                        .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                                where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                                 && ((item8["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)) ||
                                                 (item8["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                 && item8["Area"].ToString().Equals(currentArea)
                                                //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                                select item8)
                                       .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                               where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                               && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                               && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                               && item9["Area"].ToString().Equals(currentArea)
                                               && (item9["Extension"].ToString() == "Initiate")
                                               && today > (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)
                                               select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && (item9["Extension"].ToString() == "First")
                                            && today > (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && (item9["Extension"].ToString() == "First")
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today > (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && (item9["Extension"].ToString() == "Second")
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today > (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && item9["Extension"].ToString() == "Initiate"
                                            && today > Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && item9["Extension"].ToString() == "First"
                                            && today > Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && item9["Extension"].ToString() == "First"
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today > Convert.ToDateTime(item9["Extension_Date"])
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && item9["Extension"].ToString() == "Second"
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today > Convert.ToDateTime(item9["Extension_Date"])
                                            select item9)
                                .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                        where item8["MOC_x0020_Status"].ToString().Equals("Open")
                                         && ((item8["Affected_CAF"].ToString().Equals("Yes") && today > (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)) ||
                                         (item8["Affected_CAF"].ToString().Equals("") && today > (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                         && item8["Area"].ToString().Equals(currentArea)
                                         && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                        //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                        select item8)
                                                       .Select(s => new
                                                       {
                                                           CafNo = s["CAF_x0020_No"],
                                                           CAFType = s["CAFType"],
                                                           Title = s["Title"],
                                                           CRPName = s["CRP_x0020_Name"],
                                                           CurrentSection = s["Current_x0020_Section"],
                                                           InitiationDate = s["Date_x0020_Initiated"],
                                                           ConstructionDate = s["Construction_Date"],
                                                           StartUpDate = s["Est_x002e__x0020_Start_x0020_Up_"],
                                                           SubmissionDate = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6),
                                                           //SubmissionDate04E = (Convert.ToDateTime(s["ByPass_x0020_Date"])).AddMonths(6),
                                                           //SubmissionDate09 = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddYears(4),
                                                           //SubmissionDate09Ext = (Convert.ToDateTime(s["Extension_Date"])),
                                                           Remarks = s["Override_x0020_Remarks"]
                                                       });
                            ListViewCAFDetails.DataSource = getOverdueCAF;
                            ListViewCAFDetails.DataBind();
                            Logger("Area A9 OD");
                            #endregion
                        }

                        if (status.Equals("NC"))
                        {
                            #region Query NC
                            //query for NC
                            var getNCCAF = (from item in dt04.Rows.Cast<DataRow>()
                                            where item["MOC_x0020_Status"].ToString().Equals("Open") && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && today > Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])
                                            && item["Current_x0020_Section"].ToString() != "Section 5"
                                            && item["Area"].ToString().Equals(currentArea)
                                            select item)
                                      .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                              where item2["MOC_x0020_Status"].ToString().Equals("Open") && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                               && today > Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])
                                               && item2["Current_x0020_Section"].ToString() != "Section 5"
                                               && item2["Area"].ToString().Equals(currentArea)
                                              select item2)
                                       .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                               where item3["MOC_x0020_Status"].ToString().Equals("Open") && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && today > Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])
                                                && item3["Current_x0020_Section"].ToString() != "Section 5"
                                                && item3["Area"].ToString().Equals(currentArea)
                                               select item3)
                                       .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                               where item4["MOC_x0020_Status"].ToString().Equals("Open") && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && today > Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])
                                                && item4["Current_x0020_Section"].ToString() != "Section 5"
                                                && item4["Area"].ToString().Equals(currentArea)
                                               select item4)
                                       .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                               where item5["MOC_x0020_Status"].ToString().Equals("Open") && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && today > Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])
                                                && item5["Current_x0020_Section"].ToString() != "Section 5"
                                                && item5["Area"].ToString().Equals(currentArea)
                                               select item5)
                                       .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                               where item6["MOC_x0020_Status"].ToString().Equals("Open") && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && today > Convert.ToDateTime(item6["ByPass_x0020_Date"]).AddDays(90)
                                                && item6["Current_x0020_Section"].ToString() != "Section 5"
                                                && item6["Area"].ToString().Equals(currentArea)
                                               select item6)
                                       .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                               where item7["MOC_x0020_Status"].ToString().Equals("Open") && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && today > Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])
                                                && item7["Current_x0020_Section"].ToString() != "Section 5"
                                                && item7["Area"].ToString().Equals(currentArea)
                                               select item7)
                                       .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                               where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && today > Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])
                                                && item8["Current_x0020_Section"].ToString() != "Section 5"
                                                && item8["Area"].ToString().Equals(currentArea)
                                               select item8)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Extension"].ToString() == ""
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 1")
                                            || item9["Current_x0020_Section"].ToString().Equals("Section 2")
                                            || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && today > Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddDays(14)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Extension"].ToString() == ""
                                            && item9["Sect4A_x0020_AO_x0020_Approval_x"].ToString().Equals("")
                                            && !item9["Sect3_HOP_App"].ToString().Equals("")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 1")
                                            || item9["Current_x0020_Section"].ToString().Equals("Section 2")
                                            || item9["Current_x0020_Section"].ToString().Equals("Section 3")
                                            || item9["Current_x0020_Section"].ToString().Equals("Section 4"))
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && today > Convert.ToDateTime(item9["Sect3_HOP_App"]).AddDays(21)
                                            select item9)
                                .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                        where item8["MOC_x0020_Status"].ToString().Equals("Open")
                                         && today > Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])
                                         && item8["Current_x0020_Section"].ToString() != "Section 5"
                                         && item8["Area"].ToString().Equals(currentArea)
                                         && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                        select item8)
                                                           .Select(s => new
                                                           {
                                                               CafNo = s["CAF_x0020_No"],
                                                               CAFType = s["CAFType"],
                                                               Title = s["Title"],
                                                               CRPName = s["CRP_x0020_Name"],
                                                               CurrentSection = s["Current_x0020_Section"],
                                                               InitiationDate = s["Date_x0020_Initiated"],
                                                               ConstructionDate = s["Construction_Date"],
                                                               StartUpDate = s["Est_x002e__x0020_Start_x0020_Up_"],
                                                               SubmissionDate = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6),
                                                               //SubmissionDate04E = (Convert.ToDateTime(s["ByPass_x0020_Date"])).AddMonths(6),
                                                               //SubmissionDate09 = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddYears(4),
                                                               //SubmissionDate09Ext = (Convert.ToDateTime(s["Extension_Date"])),
                                                               Remarks = s["Override_x0020_Remarks"]
                                                           });
                            ListViewCAFDetails.DataSource = getNCCAF;
                            ListViewCAFDetails.DataBind();
                            Logger("Area A9 NC");
                            #endregion
                        }

                        if (status.Equals("GTO"))
                        {
                            #region Query Going to Overdue 2 weeks
                            //query for Going to Overdue 2 weeks
                            var getGoingToOverdue2weekCAF = (from item in dt04.Rows.Cast<DataRow>()
                                                             where item["MOC_x0020_Status"].ToString().Equals("Open")
                                                             && ((item["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                                             && today <= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                             || (item["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                                             && today <= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                             && item["Area"].ToString().Equals(currentArea)
                                                             && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                                             select item)
                                        .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                                where item2["MOC_x0020_Status"].ToString().Equals("Open") && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && ((item2["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                || (item2["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                && item2["Area"].ToString().Equals(currentArea)
                                                select item2)
                                        .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                                where item3["MOC_x0020_Status"].ToString().Equals("Open") && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && ((item3["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                || (item3["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                && item3["Area"].ToString().Equals(currentArea)
                                                select item3)
                                        .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                                where item4["MOC_x0020_Status"].ToString().Equals("Open") && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && ((item4["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                || (item4["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                && item4["Area"].ToString().Equals(currentArea)
                                                select item4)
                                        .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                                where item5["MOC_x0020_Status"].ToString().Equals("Open") && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && ((item5["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                || (item5["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                && item5["Area"].ToString().Equals(currentArea)
                                                select item5)
                                        .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                                where item6["MOC_x0020_Status"].ToString().Equals("Open")
                                                 && today >= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddMonths(6).AddDays(-14)
                                                 && today <= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddMonths(6)
                                                 && item6["Area"].ToString().Equals(currentArea)
                                                 && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                                //&& item6["Current_x0020_Section"].ToString() != "Section 5"
                                                select item6)
                                        .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                                where item7["MOC_x0020_Status"].ToString().Equals("Open") && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && ((item7["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                || (item7["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                && item7["Area"].ToString().Equals(currentArea)
                                                select item7)
                                        .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                                where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && ((item8["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                || (item8["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                && item8["Area"].ToString().Equals(currentArea)
                                                select item8)
                                        .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                                where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                                && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                                && item9["Area"].ToString().Equals(currentArea)
                                                && (item9["Extension"].ToString() == "Initiate")
                                                && today >= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)
                                                select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && (item9["Extension"].ToString() == "First")
                                           && today >= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "First")
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && (item9["Extension"].ToString() == "Second")
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && item9["Extension"].ToString() == "Initiate"
                                            && today >= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4).AddDays(-14)
                                            && today <= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && item9["Extension"].ToString() == "First"
                                            && today >= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4).AddDays(-14)
                                            && today <= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && item9["Extension"].ToString() == "First"
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && item9["Extension"].ToString() == "Second"
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddDays(-14)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                        where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                        && ((item8["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6).AddDays(-14)
                                        && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                        || (item8["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3).AddDays(-14)
                                        && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                        && item8["Area"].ToString().Equals(currentArea)
                                        && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                        select item8)
                                        .Select(s => new
                                        {
                                            CafNo = s["CAF_x0020_No"],
                                            CAFType = s["CAFType"],
                                            Title = s["Title"],
                                            CRPName = s["CRP_x0020_Name"],
                                            CurrentSection = s["Current_x0020_Section"],
                                            InitiationDate = s["Date_x0020_Initiated"],
                                            ConstructionDate = s["Construction_Date"],
                                            StartUpDate = s["Est_x002e__x0020_Start_x0020_Up_"],
                                            SubmissionDate = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6),
                                            //SubmissionDate04E = (Convert.ToDateTime(s["ByPass_x0020_Date"])).AddMonths(6),
                                            //SubmissionDate09 = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddYears(4),
                                            //SubmissionDate09Ext = (Convert.ToDateTime(s["Extension_Date"])),
                                            Remarks = s["Override_x0020_Remarks"]
                                        });
                            ListViewCAFDetails.DataSource = getGoingToOverdue2weekCAF;
                            ListViewCAFDetails.DataBind();
                            Logger("Area A9 OD 2W");
                            #endregion
                        }

                        if (status.Equals("GTNC"))
                        {
                            #region Going to NC 2 weeks
                            //query for Going to NC 2 weeks
                            var getGoingtoNC2weeks = (from item in dt04.Rows.Cast<DataRow>()
                                                      where item["MOC_x0020_Status"].ToString().Equals("Open")
                                                      && today >= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                                      && today <= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"]))
                                                      && item["Current_x0020_Section"].ToString() != "Section 5"
                                                      && item["Area"].ToString().Equals(currentArea)
                                                      && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                                      select item)
                                       .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                               where item2["MOC_x0020_Status"].ToString().Equals("Open")
                                                && today >= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"]))
                                                && item2["Current_x0020_Section"].ToString() != "Section 5"
                                                && item2["Area"].ToString().Equals(currentArea)
                                                && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item2)
                                       .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                               where item3["MOC_x0020_Status"].ToString().Equals("Open")
                                                && today >= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"]))
                                                && item3["Current_x0020_Section"].ToString() != "Section 5"
                                                && item3["Area"].ToString().Equals(currentArea)
                                                && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item3)
                                       .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                               where item4["MOC_x0020_Status"].ToString().Equals("Open")
                                                && today >= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"]))
                                                && item4["Current_x0020_Section"].ToString() != "Section 5"
                                                && item4["Area"].ToString().Equals(currentArea)
                                                && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item4)
                                       .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                               where item5["MOC_x0020_Status"].ToString().Equals("Open")
                                                && today >= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"]))
                                                && item5["Current_x0020_Section"].ToString() != "Section 5"
                                                && item5["Area"].ToString().Equals(currentArea)
                                                && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item5)
                                       .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                               where item6["MOC_x0020_Status"].ToString().Equals("Open")
                                                && today >= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddDays(90).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddDays(90)
                                                && item6["Current_x0020_Section"].ToString() != "Section 5"
                                                && item6["Area"].ToString().Equals(currentArea)
                                                && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item6)
                                       .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                               where item7["MOC_x0020_Status"].ToString().Equals("Open")
                                                && today >= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"]))
                                                && item7["Current_x0020_Section"].ToString() != "Section 5"
                                                && item7["Area"].ToString().Equals(currentArea)
                                                && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item7)
                                       .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                               where item8["MOC_x0020_Status"].ToString().Equals("Open")
                                                && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                                && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"]))
                                                && item8["Current_x0020_Section"].ToString() != "Section 5"
                                                && item8["Area"].ToString().Equals(currentArea)
                                                && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item8)
                                       .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                               where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                               && item9["Extension"].ToString() == ""
                                               && (item9["Current_x0020_Section"].ToString().Equals("Section 1")
                                               || item9["Current_x0020_Section"].ToString().Equals("Section 2")
                                               || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                               && item9["Area"].ToString().Equals(currentArea)
                                               && today >= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])
                                               && today <= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(14)
                                               select item9)
                                .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                        where item8["MOC_x0020_Status"].ToString().Equals("Open")
                                         && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(-14)
                                         && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"]))
                                         && item8["Current_x0020_Section"].ToString() != "Section 5"
                                         && item8["Area"].ToString().Equals(currentArea)
                                         && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                        select item8)
                                            .Select(s => new
                                            {
                                                CafNo = s["CAF_x0020_No"],
                                                CAFType = s["CAFType"],
                                                Title = s["Title"],
                                                CRPName = s["CRP_x0020_Name"],
                                                CurrentSection = s["Current_x0020_Section"],
                                                InitiationDate = s["Date_x0020_Initiated"],
                                                ConstructionDate = s["Construction_Date"],
                                                StartUpDate = s["Est_x002e__x0020_Start_x0020_Up_"],
                                                SubmissionDate = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6),
                                                //SubmissionDate04E = (Convert.ToDateTime(s["ByPass_x0020_Date"])).AddMonths(6),
                                                //SubmissionDate09 = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddYears(4),
                                                //SubmissionDate09Ext = (Convert.ToDateTime(s["Extension_Date"])),
                                                Remarks = s["Override_x0020_Remarks"]
                                            });
                            ListViewCAFDetails.DataSource = getGoingtoNC2weeks;
                            ListViewCAFDetails.DataBind();
                            Logger("Area A9 NC 2W");
                            #endregion
                        }

                        if (status.Equals("GTO_1month"))
                        {
                            #region Query Going To Overdue 1 month
                            //query for Going to Overdue 1 mth
                            var getGoingToOverdue1mth = (from item in dt04.Rows.Cast<DataRow>()
                                                         where item["MOC_x0020_Status"].ToString().Equals("Open")
                                                         && ((item["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                                         && today <= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                         || (item["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                                         && today <= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                         && item["Area"].ToString().Equals(currentArea)
                                                         && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                                         //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                                         select item)
                                        .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                                where item2["MOC_x0020_Status"].ToString().Equals("Open") && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && ((item2["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                                && today <= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                || (item2["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                                && today <= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                && item2["Area"].ToString().Equals(currentArea)
                                                //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                                select item2)
                                        .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                                where item3["MOC_x0020_Status"].ToString().Equals("Open") && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && ((item3["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                                && today <= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                || (item3["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                                && today <= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                && item3["Area"].ToString().Equals(currentArea)
                                                //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                                select item3)
                                        .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                                where item4["MOC_x0020_Status"].ToString().Equals("Open") && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && ((item4["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                                && today <= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                || (item4["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                                && today <= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                && item4["Area"].ToString().Equals(currentArea)
                                                //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                                select item4)
                                        .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                                where item5["MOC_x0020_Status"].ToString().Equals("Open") && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && ((item5["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                                && today <= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                || (item5["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                                && today <= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                && item5["Area"].ToString().Equals(currentArea)
                                                //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                                select item5)
                                        .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                                where item6["MOC_x0020_Status"].ToString().Equals("Open")
                                                 && today >= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddMonths(5)
                                                 && today <= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddMonths(6)
                                                 && item6["Area"].ToString().Equals(currentArea)
                                                 && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                                //&& item6["Current_x0020_Section"].ToString() != "Section 5"
                                                select item6)
                                        .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                                where item7["MOC_x0020_Status"].ToString().Equals("Open") && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && ((item7["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                                && today <= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                || (item7["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                                && today <= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                && item7["Area"].ToString().Equals(currentArea)
                                                //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                                select item7)
                                        .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                                where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                                && ((item8["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                                && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                                || (item8["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                                && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                                && item8["Area"].ToString().Equals(currentArea)
                                                //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                                select item8)
                                          .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                                  where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                                  && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                                  && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                                  && item9["Area"].ToString().Equals(currentArea)
                                                  && (item9["Extension"].ToString() == "Initiate")
                                                  && today >= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                                  && today <= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)
                                                  select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && (item9["Extension"].ToString() == "First")
                                           && today >= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                            && today <= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && (item9["Extension"].ToString() == "First")
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && (item9["Extension"].ToString() == "Second")
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && item9["Extension"].ToString() == "Initiate"
                                            && today >= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(3)
                                            && today <= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && item9["Extension"].ToString() == "First"
                                            && today >= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(3)
                                            && today <= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Current_x0020_Section"].ToString().Equals("Section 4")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && item9["Extension"].ToString() == "First"
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && (item9["Current_x0020_Section"].ToString().Equals("Section 2") || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && item9["Area"].ToString().Equals(currentArea)
                                            && item9["Extension"].ToString() == "Second"
                                            && !(item9["Extension_Date"].ToString().Equals(""))
                                            && today >= (Convert.ToDateTime(item9["Extension_Date"])).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item9["Extension_Date"]))
                                            select item9)
                                .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                        where item8["MOC_x0020_Status"].ToString().Equals("Open") && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                        && ((item8["Affected_CAF"].ToString().Equals("Yes") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(5)
                                        && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6))
                                        || (item8["Affected_CAF"].ToString().Equals("") && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(2)
                                        && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(3)))
                                        && item8["Area"].ToString().Equals(currentArea)
                                        && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                        //&& item["Current_x0020_Section"].ToString() != "Section 5"
                                        select item8)
                                                .Select(s => new
                                                {
                                                    CafNo = s["CAF_x0020_No"],
                                                    CAFType = s["CAFType"],
                                                    Title = s["Title"],
                                                    CRPName = s["CRP_x0020_Name"],
                                                    CurrentSection = s["Current_x0020_Section"],
                                                    InitiationDate = s["Date_x0020_Initiated"],
                                                    ConstructionDate = s["Construction_Date"],
                                                    StartUpDate = s["Est_x002e__x0020_Start_x0020_Up_"],
                                                    SubmissionDate = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6),
                                                    //SubmissionDate04E = (Convert.ToDateTime(s["ByPass_x0020_Date"])).AddMonths(6),
                                                    //SubmissionDate09 = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddYears(4),
                                                    //SubmissionDate09Ext = (Convert.ToDateTime(s["Extension_Date"])),
                                                    Remarks = s["Override_x0020_Remarks"]
                                                });
                            ListViewCAFDetails.DataSource = getGoingToOverdue1mth;
                            ListViewCAFDetails.DataBind();
                            Logger("Area A9 OD 1M");
                            #endregion
                        }

                        if (status.Equals("GTNC_1month"))
                        {
                            #region Query To NC 1 Month
                            //query for Going to NC 1 mth
                            var getGointToNC1mth = (from item in dt04.Rows.Cast<DataRow>()
                                                    where item["MOC_x0020_Status"].ToString().Equals("Open")
                                                    && today >= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                                    && today <= (Convert.ToDateTime(item["Est_x002e__x0020_Start_x0020_Up_"]))
                                                    && item["Current_x0020_Section"].ToString() != "Section 5"
                                                    && item["Area"].ToString().Equals(currentArea)
                                                    && item["DATA_x0020_TYPE"].ToString().Equals("new")
                                                    select item)
                                       .Concat(from item2 in dt04A.Rows.Cast<DataRow>()
                                               where item2["MOC_x0020_Status"].ToString().Equals("Open")
                                                && today >= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                                && today <= (Convert.ToDateTime(item2["Est_x002e__x0020_Start_x0020_Up_"]))
                                                && item2["Current_x0020_Section"].ToString() != "Section 5"
                                                && item2["Area"].ToString().Equals(currentArea)
                                                && item2["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item2)
                                       .Concat(from item3 in dt04B.Rows.Cast<DataRow>()
                                               where item3["MOC_x0020_Status"].ToString().Equals("Open")
                                                && today >= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                                && today <= (Convert.ToDateTime(item3["Est_x002e__x0020_Start_x0020_Up_"]))
                                                && item3["Current_x0020_Section"].ToString() != "Section 5"
                                                && item3["Area"].ToString().Equals(currentArea)
                                                && item3["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item3)
                                       .Concat(from item4 in dt04C.Rows.Cast<DataRow>()
                                               where item4["MOC_x0020_Status"].ToString().Equals("Open")
                                                && today >= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                                && today <= (Convert.ToDateTime(item4["Est_x002e__x0020_Start_x0020_Up_"]))
                                                && item4["Current_x0020_Section"].ToString() != "Section 5"
                                                && item4["Area"].ToString().Equals(currentArea)
                                                && item4["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item4)
                                       .Concat(from item5 in dt04D.Rows.Cast<DataRow>()
                                               where item5["MOC_x0020_Status"].ToString().Equals("Open")
                                                && today >= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                                && today <= (Convert.ToDateTime(item5["Est_x002e__x0020_Start_x0020_Up_"]))
                                                && item5["Current_x0020_Section"].ToString() != "Section 5"
                                                && item5["Area"].ToString().Equals(currentArea)
                                                && item5["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item5)
                                       .Concat(from item6 in dt04E.Rows.Cast<DataRow>()
                                               where item6["MOC_x0020_Status"].ToString().Equals("Open")
                                                && today >= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddDays(90).AddMonths(-1)
                                                && today <= (Convert.ToDateTime(item6["ByPass_x0020_Date"])).AddDays(90)
                                                && item6["Current_x0020_Section"].ToString() != "Section 5"
                                                && item6["Area"].ToString().Equals(currentArea)
                                                && item6["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item6)
                                       .Concat(from item7 in dt04G.Rows.Cast<DataRow>()
                                               where item7["MOC_x0020_Status"].ToString().Equals("Open")
                                                && today >= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                                && today <= (Convert.ToDateTime(item7["Est_x002e__x0020_Start_x0020_Up_"]))
                                                && item7["Current_x0020_Section"].ToString() != "Section 5"
                                                && item7["Area"].ToString().Equals(currentArea)
                                                && item7["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item7)
                                       .Concat(from item8 in dt04J.Rows.Cast<DataRow>()
                                               where item8["MOC_x0020_Status"].ToString().Equals("Open")
                                                && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                                && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"]))
                                                && item8["Current_x0020_Section"].ToString() != "Section 5"
                                                && item8["Area"].ToString().Equals(currentArea)
                                                && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                               select item8)
                                       .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                               where item9["MOC_x0020_Status"].ToString().Equals("Open") && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                               && item9["Extension"].ToString() == ""
                                               && (item9["Current_x0020_Section"].ToString().Equals("Section 1")
                                               || item9["Current_x0020_Section"].ToString().Equals("Section 2")
                                               || item9["Current_x0020_Section"].ToString().Equals("Section 3"))
                                               && item9["Area"].ToString().Equals(currentArea)
                                               && today >= Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"]).AddDays(14).AddMonths(-1)
                                               && today <= (Convert.ToDateTime(item9["Est_x002e__x0020_Start_x0020_Up_"])).AddDays(14)
                                               select item9)
                                .Concat(from item8 in dt04L.Rows.Cast<DataRow>()
                                        where item8["MOC_x0020_Status"].ToString().Equals("Open")
                                         && today >= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(-1)
                                         && today <= (Convert.ToDateTime(item8["Est_x002e__x0020_Start_x0020_Up_"]))
                                         && item8["Current_x0020_Section"].ToString() != "Section 5"
                                         && item8["Area"].ToString().Equals(currentArea)
                                         && item8["DATA_x0020_TYPE"].ToString().Equals("new")
                                        select item8)
                                                    .Select(s => new
                                                    {
                                                        CafNo = s["CAF_x0020_No"],
                                                        CAFType = s["CAFType"],
                                                        Title = s["Title"],
                                                        CRPName = s["CRP_x0020_Name"],
                                                        CurrentSection = s["Current_x0020_Section"],
                                                        InitiationDate = s["Date_x0020_Initiated"],
                                                        ConstructionDate = s["Construction_Date"],
                                                        StartUpDate = s["Est_x002e__x0020_Start_x0020_Up_"],
                                                        SubmissionDate = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6),
                                                        //SubmissionDate04E = (Convert.ToDateTime(s["ByPass_x0020_Date"])).AddMonths(6),
                                                        //SubmissionDate09 = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddYears(4),
                                                        //SubmissionDate09Ext = (Convert.ToDateTime(s["Extension_Date"])),
                                                        Remarks = s["Override_x0020_Remarks"]
                                                    });
                            ListViewCAFDetails.DataSource = getGointToNC1mth;
                            ListViewCAFDetails.DataBind();
                            Logger("Area A9 NC 1M");
                            #endregion

                        }

                        if (status.Equals("Ext_OLS"))
                        {
                            #region Query Ext OLS
                            //query for Going to NC 1 mth
                            var getExtOLS = (from item9 in dt09.Rows.Cast<DataRow>()
                                             where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                             && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                             && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                             && (item9["Extension"].ToString() == "Initiate" || item9["Extension"].ToString() == "")
                                             && !item9["Actual_x0020_Date"].ToString().Equals("")
                                             && today > (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(6)
                                             select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "First")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today > (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(12)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "Second")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today > (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(18)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && (item9["Extension"].ToString() == "Initiate" || item9["Extension"].ToString() == "")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today > (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && (item9["Extension"].ToString() == "First")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today > (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(5)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && (item9["Extension"].ToString() == "Second")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today > (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(6)
                                            select item9)
                                                    .Select(s => new
                                                    {
                                                        CafNo = s["CAF_x0020_No"],
                                                        CAFType = s["CAFType"],
                                                        Title = s["Title"],
                                                        CRPName = s["CRP_x0020_Name"],
                                                        CurrentSection = s["Current_x0020_Section"],
                                                        InitiationDate = s["Date_x0020_Initiated"],
                                                        ConstructionDate = s["Construction_Date"],
                                                        StartUpDate = s["Est_x002e__x0020_Start_x0020_Up_"],
                                                        SubmissionDate = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6),
                                                        //SubmissionDate04E = (Convert.ToDateTime(s["ByPass_x0020_Date"])).AddMonths(6),
                                                        //SubmissionDate09 = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddYears(4),
                                                        //SubmissionDate09Ext = (Convert.ToDateTime(s["Extension_Date"])),
                                                        Remarks = s["Override_x0020_Remarks"]
                                                    });
                            ListViewCAFDetails.DataSource = getExtOLS;
                            ListViewCAFDetails.DataBind();
                            Logger("Area A9 Ext OLS");
                            #endregion

                        }

                        if (status.Equals("Ext_OLS_1m"))
                        {
                            #region Query Ext OLS 1m
                            //query for Going to NC 1 mth
                            var getExtOLS1M = (from item9 in dt09.Rows.Cast<DataRow>()
                                               where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                               && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                               && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                               && (item9["Extension"].ToString() == "Initiate" || item9["Extension"].ToString() == "")
                                               && !item9["Actual_x0020_Date"].ToString().Equals("")
                                               && today >= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(5)
                                               && today <= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(6)
                                               select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "First")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today >= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(11)
                                            && today <= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(12)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Non Standard Repair")
                                            && (item9["Extension"].ToString() == "Second")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today >= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(17)
                                            && today <= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddMonths(18)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && (item9["Extension"].ToString() == "Initiate" || item9["Extension"].ToString() == "")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today >= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(4).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(4)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && (item9["Extension"].ToString() == "First")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today >= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(5).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(5)
                                            select item9)
                                    .Concat(from item9 in dt09.Rows.Cast<DataRow>()
                                            where item9["MOC_x0020_Status"].ToString().Equals("Open")
                                            && item9["DATA_x0020_TYPE"].ToString().Equals("new")
                                            && item9["Type_x0020_of_x0020_Repair"].ToString().Equals("Temporary Repair")
                                            && (item9["Extension"].ToString() == "Second")
                                            && !item9["Actual_x0020_Date"].ToString().Equals("")
                                            && today >= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(6).AddMonths(-1)
                                            && today <= (Convert.ToDateTime(item9["Actual_x0020_Date"])).AddYears(6)
                                            select item9)
                                                    .Select(s => new
                                                    {
                                                        CafNo = s["CAF_x0020_No"],
                                                        CAFType = s["CAFType"],
                                                        Title = s["Title"],
                                                        CRPName = s["CRP_x0020_Name"],
                                                        CurrentSection = s["Current_x0020_Section"],
                                                        InitiationDate = s["Date_x0020_Initiated"],
                                                        ConstructionDate = s["Construction_Date"],
                                                        StartUpDate = s["Est_x002e__x0020_Start_x0020_Up_"],
                                                        SubmissionDate = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddMonths(6),
                                                        //SubmissionDate04E = (Convert.ToDateTime(s["ByPass_x0020_Date"])).AddMonths(6),
                                                        //SubmissionDate09 = (Convert.ToDateTime(s["Est_x002e__x0020_Start_x0020_Up_"])).AddYears(4),
                                                        //SubmissionDate09Ext = (Convert.ToDateTime(s["Extension_Date"])),
                                                        Remarks = s["Override_x0020_Remarks"]
                                                    });
                            ListViewCAFDetails.DataSource = getExtOLS1M;
                            ListViewCAFDetails.DataBind();
                            Logger("Area A9 Ext OLS 1M");
                            #endregion

                        }
                    }
                    catch (Exception ex) {
                        Logger("Error catch: "+ex.ToString());
                    }
                    /**/
                    /*tableEmocCAF = new DataTable();
                    tableEmocCAF.Columns.Add("CAF_Number");
                    tableEmocCAF.Columns.Add("CAF_Type");
                    tableEmocCAF.Columns.Add("CAF_Title");
                    tableEmocCAF.Columns.Add("CRP");
                    tableEmocCAF.Columns.Add("Current_Status");
                    tableEmocCAF.Columns.Add("Initiation_Date");
                    tableEmocCAF.Columns.Add("Construction_Date");
                    tableEmocCAF.Columns.Add("StartUp_Date");
                    tableEmocCAF.Columns.Add("Submission_Date");*/


                    //ListViewCAF.DataSource = tableEmocCAF;
                    //ListViewCAFDetails.DataSource
                    //ListViewCAFDetails.DataBind();
                }


            }
        

            //oWeb.Dispose();
            //oSite.Dispose();
        }

        protected void btnBack_Click(object sender, EventArgs e)
        {
            ListViewCAFDetails.Visible = false;
            ListViewCAF.Visible = true;
            btnBack.Visible = false;
        }

        public static void Logger(string lines)
        {
            string path = "C:/Log/";
            VerifyDir(path);
            string fileName = DateTime.Now.Day.ToString() + DateTime.Now.Month.ToString() + DateTime.Now.Year.ToString() + "_Logs.txt";
            try
            {
                System.IO.StreamWriter file = new System.IO.StreamWriter(path + fileName, true);
                file.WriteLine(DateTime.Now.ToString() + ": " + lines);
                file.Close();
            }
            catch (Exception) { }
        }

        public static void VerifyDir(string path)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(path);
                if (!dir.Exists)
                {
                    dir.Create();
                }
            }
            catch { }
        }
    }
}
