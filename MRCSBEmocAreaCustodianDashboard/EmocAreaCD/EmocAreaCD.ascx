<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %> 
<%@ Register Tagprefix="WebPartPages" Namespace="Microsoft.SharePoint.WebPartPages" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="EmocAreaCD.ascx.cs" Inherits="MRCSBEmocAreaCustodianDashboard.VisualWebPart1.VisualWebPart1" %>
<style type="text/css">
    .DashMOC_Heading {
        background-color: #01bbb2;
        color: white !important;
    }
    .DashMOC_header {
        color: black;
        font-family: Verdana !important;
        font-size: 14px !important;
        font-weight: bold !important;
        padding: 5px;
        text-align: center !important;
        text-transform: uppercase;
        border:1px solid black;
    }
    .DashMOC_header_Area {
        background-color: #5b9bd5;
        width:20%;
    }
    .DashMOC_header_Open {
        background-color: #70ad47;
        width:10%;
    }
    .DashMOC_header_Open_Type1 {
        background-color: #93e05f;
        width:10%;
    }
    .DashMOC_header_Open_Type2 {
        background-color: #81c753;
        width:10%;
    }
    .DashMOC_header_Close {
        background-color: #afabab;
        width:10%
    }
    .DashMOC_header_Overdue {
        background-color: #ffc000;
        width:10%
    }
    .DashMOC_header_NC {
        background-color: #ff0000;
        width:10%
    }
    .DashMOC_header_Overdue2w {
        background-color: #ffd966;
        width:10%
    }
    .DashMOC_header_NC2w {
        background-color: #ff5050;
        width:10%
    }
    .DashMOC_header_Overdue1m {
        background-color: #ffe699;
        width:10%
    }
    .DashMOC_header_NC1m {
        background-color: #ff7c80;
        width:10%
    }
    .DashMOC_header_ExtOLS {
        background-color: #ffc000;
        width:10%
    }
    .DashMOC_header_ExtOLS1m {
        background-color: #ff7c80;
        width:10%
    }

    .DashMOC_content_Area {
        color: black;
        font-size: 14px;
        text-align: center;
        background-color: #5b9bd5;
        border:1px solid black;
        padding: 5px;
        width:20%;
        
    }
    .DashMOC_content_Open {
        color: black;
        font-size: 14px;
        text-align: center;
        background-color: #70ad47;
        border:1px solid black;
        padding: 5px;
        width:10%;
    }
    .DashMOC_content_Open_Type1 {
        color: black;
        font-size: 14px;
        text-align: center;
        background-color: #93e05f;
        border:1px solid black;
        padding: 5px;
        width:5%;
    }
    .DashMOC_content_Open_Type2 {
        color: black;
        font-size: 14px;
        text-align: center;
        background-color: #81c753;
        border:1px solid black;
        padding: 5px;
        width:5%;
    }
    .DashMOC_content_Close {
        color: black;
        font-size: 14px;
        text-align: center;
        background-color: #afabab;
        border:1px solid black;
        padding: 5px;
        width:10%
    }
    .DashMOC_content_Overdue {
        color: black;
        font-size: 14px;
        text-align: center;
        background-color: #ffc000;
        border:1px solid black;
        padding: 5px;
        width:10%
    }
    .DashMOC_content_NC {
        color: black;
        font-size: 14px;
        text-align: center;
        background-color: #ff0000;
        border:1px solid black;
        padding: 5px;
        width:10%
    }
    .DashMOC_content_Overdue2w {
        color: black;
        font-size: 14px;
        text-align: center;
        background-color: #ffd966;
        border:1px solid black;
        padding: 5px;
        width:10%
    }
    .DashMOC_content_NC2w {
        color: black;
        font-size: 14px;
        text-align: center;
        background-color: #ff5050;
        border:1px solid black;
        padding: 5px;
        width:10%
    }
    .DashMOC_content_Overdue1m {
        color: black;
        font-size: 14px;
        text-align: center;
        background-color: #ffe699;
        border:1px solid black;
        padding: 5px;
        width:10%
    }
    .DashMOC_content_NC1m {
        color: black;
        font-size: 14px;
        text-align: center;
        background-color: #ff7c80;
        border:1px solid black;
        padding: 5px;
        width:10%
    }
    .DashMOC_content_ExtOLS {
        color: black;
        font-size: 14px;
        text-align: center;
        background-color: #ffc000;
        border:1px solid black;
        padding: 5px;
        width:10%
    }
    .DashMOC_content_ExtOLS1m {
        color: black;
        font-size: 14px;
        text-align: center;
        background-color: #ff7c80;
        border:1px solid black;
        padding: 5px;
        width:10%
    }

    .DashMOC_Link a:link {
    color: black;
    background-color: transparent;
    text-decoration: none;
        font-weight: bold !important;
    }
    .DashMOC_Visit a:visited {
        color: black;
        background-color: transparent;
        text-decoration: none;
    }
    .DashMOC_Hover a:hover {
        color: black;
        background-color: #01bbb2;
        text-decoration: underline;
        font-weight: bold !important;
    }
</style>
<table border="0" style="width: 100%;">
    <tr>
        <td>
            <asp:Label ID="Label1" runat="server" Text="As of : "></asp:Label>
            <asp:Label ID="lblDate" runat="server" Text=""></asp:Label>
        </td>
    </tr>
    <tr>
        <td class="DashMOC_header DashMOC_Heading">EMOC Custodian Dashboard</td>
    </tr>

    <tr>
        <td>
        <asp:ListView ID="ListViewCAF" runat="server" OnItemCommand="ListViewCAF_ItemCommand">
            <LayoutTemplate>
                <table border="0" style="width: 100%" >
                    <tr>
                        <td rowspan="2" class="DashMOC_header DashMOC_header_Area">Area
                        </td>
                        <td colspan="2" class="DashMOC_header DashMOC_header_Open">Open
                        </td>
                        <td rowspan="2" class="DashMOC_header DashMOC_header_Close">Close
                        </td>
                        <td rowspan="2" class="DashMOC_header DashMOC_header_Overdue">Overdue
                        </td>
                        <td rowspan="2" class="DashMOC_header DashMOC_header_NC">Non-Compliance
                        </td>
                        <td rowspan="2" class="DashMOC_header DashMOC_header_Overdue2w">Going to Overdue (2 weeks)
                        </td>
                        <td rowspan="2" class="DashMOC_header DashMOC_header_NC2w">Going to Non Compliance (2 weeks)
                        </td>
                        <td rowspan="2" class="DashMOC_header DashMOC_header_Overdue1m">Going to Overdue (1 month)
                        </td>
                        <td rowspan="2" class="DashMOC_header DashMOC_header_NC1m">Going to Non Compliance (1 month)
                        </td>
                        <td rowspan="2" class="DashMOC_header DashMOC_header_ExtOLS">Overdue Extension OLS
                        </td>
                        <td rowspan="2" class="DashMOC_header DashMOC_header_ExtOLS1m">Going to Overdue Extension OLS (1 month)
                        </td>
                    </tr>
                    <tr>
                        <td class="DashMOC_header DashMOC_header_Open_Type1">Type 1</td>
                        <td class="DashMOC_header DashMOC_header_Open_Type2">Type 2</td>
                    </tr>
                    <tr runat="server" id="itemPlaceholder">
                    </tr>
                </table>
            </LayoutTemplate>
            <ItemTemplate>
                <tr>
                    <td class="DashMOC_content_Area DashMOC_Link DashMOC_Hover DashMOC_Visit">
                        <asp:Label ID="lblArea" runat="server" Text='<%# Eval("Area") %>' Visible="True"></asp:Label>
                    </td>
                    <td class="DashMOC_content_Open_Type1 DashMOC_Link DashMOC_Hover DashMOC_Visit">
                         <%--<asp:Label ID="lblOpen" runat="server" Text='<%# Eval("Open") %>' Visible="false" CssClass="textAlign"></asp:Label>--%>
                        <asp:LinkButton ID="linkOpen" CommandName="Open" runat="server" Text='<%# Eval("Open") %>'></asp:LinkButton>
                    </td>
                    <td class="DashMOC_content_Open_Type2 DashMOC_Link DashMOC_Hover DashMOC_Visit">
                         <%--<asp:Label ID="lblOpen" runat="server" Text='<%# Eval("Open") %>' Visible="false" CssClass="textAlign"></asp:Label>--%>
                        <asp:LinkButton ID="linkOpen2" CommandName="Open2" runat="server" Text='<%# Eval("Open2") %>'></asp:LinkButton>
                    </td>
                    <td class="DashMOC_content_Close DashMOC_Link DashMOC_Hover DashMOC_Visit">
                        <asp:Label ID="lblClose" runat="server" Text='<%# Eval("Close") %>' Visible="false" CssClass="textAlign"></asp:Label>
                        <asp:LinkButton ID="linkClose" CommandName="Closed" runat="server" Text='<%# Eval("Close") %>'></asp:LinkButton>
                    </td>
                     <td class="DashMOC_content_Overdue DashMOC_Link DashMOC_Hover DashMOC_Visit">
                        <asp:Label ID="lblOverdue" runat="server" Text='<%# Eval("Overdue") %>' Visible="false" CssClass="textAlign"></asp:Label>
                         <asp:LinkButton ID="linkOverdue" CommandName="Overdue" runat="server" Text='<%# Eval("Overdue") %>'></asp:LinkButton>
                    </td>
                     <td class="DashMOC_content_NC DashMOC_Link DashMOC_Hover DashMOC_Visit">
                        <asp:Label ID="lblNC" runat="server" Text='<%# Eval("Non_Compliance") %>' Visible="false" CssClass="textAlign"></asp:Label>
                         <asp:LinkButton ID="linkNC" CommandName="NC" runat="server" Text='<%# Eval("Non_Compliance") %>'></asp:LinkButton>
                    </td>
                     <td class="DashMOC_content_Overdue2w DashMOC_Link DashMOC_Hover DashMOC_Visit">
                        <asp:Label ID="lblGTOverdue2Weeks" runat="server" Text='<%# Eval("G_T_Overdue_2Weeks") %>' Visible="false" CssClass="textAlign"></asp:Label>
                        <asp:LinkButton ID="linkGTOverdue2Weeks" CommandName="GTO" runat="server" Text='<%# Eval("G_T_Overdue_2Weeks") %>'></asp:LinkButton>
                     </td>
                     <td class="DashMOC_content_NC2w DashMOC_Link DashMOC_Hover DashMOC_Visit">
                        <asp:Label ID="lblGTNC2Weeks" runat="server" Text='<%# Eval("G_T_NC_2Weeks") %>' Visible="false" CssClass="textAlign"></asp:Label>
                        <asp:LinkButton ID="linkGTNC2weeks" CommandName="GTNC" runat="server" Text='<%# Eval("G_T_NC_2Weeks") %>'></asp:LinkButton>
                     </td>
                     <td class="DashMOC_content_Overdue1m DashMOC_Link DashMOC_Hover DashMOC_Visit">
                        <asp:Label ID="lblGTOverdue1mth" runat="server" Text='<%# Eval("G_T_Overdue_1mth") %>' Visible="false" CssClass="textAlign"></asp:Label>
                        <asp:LinkButton ID="linkGTOverdue1mth" CommandName="GTO_1month" runat="server" Text='<%# Eval("G_T_Overdue_1mth") %>'></asp:LinkButton>
                     </td>
                     <td class="DashMOC_content_NC1m DashMOC_Link DashMOC_Hover DashMOC_Visit">
                        <asp:Label ID="lblGTNC1mth" runat="server" Text='<%# Eval("G_T_NC_1mth") %>' Visible="false" CssClass="textAlign"></asp:Label>
                        <asp:LinkButton ID="linknkGTNC1mth" CommandName="GTNC_1month" runat="server" Text='<%# Eval("G_T_NC_1mth") %>'></asp:LinkButton>
                     </td>
                    <td class="DashMOC_content_ExtOLS DashMOC_Link DashMOC_Hover DashMOC_Visit">
                        <asp:Label ID="lblExtOLS" runat="server" Text='<%# Eval("Ext_OLS") %>' Visible="false" CssClass="textAlign"></asp:Label>
                        <asp:LinkButton ID="linkExt_OLS" CommandName="Ext_OLS" runat="server" Text='<%# Eval("Ext_OLS") %>'></asp:LinkButton>
                     </td>
                    <td class="DashMOC_content_ExtOLS1m DashMOC_Link DashMOC_Hover DashMOC_Visit">
                        <asp:Label ID="lblExtOLS1M" runat="server" Text='<%# Eval("Ext_OLS_1m") %>' Visible="false" CssClass="textAlign"></asp:Label>
                        <asp:LinkButton ID="linkExt_OLS_1m" CommandName="Ext_OLS_1m" runat="server" Text='<%# Eval("Ext_OLS_1m") %>'></asp:LinkButton>
                     </td>
                </tr>
            </ItemTemplate>
        </asp:ListView>
       </td>
    </tr>
    <tr>
        <td>
            <asp:Button ID="btnBack" runat="server" Text="Back" OnClick="btnBack_Click" />
            <br />
            <asp:ListView ID="ListViewCAFDetails" runat="server">
            <LayoutTemplate>
                <table border="1" style="width: 100%" >
                    <tr>
                        <td class="DashMOC_header DashMOC_Heading">CAF Number</td>
                        <td class="DashMOC_header DashMOC_Heading">CAF Type</td>
                        <td class="DashMOC_header DashMOC_Heading">CAF Title</td>
                        <td class="DashMOC_header DashMOC_Heading">CRP</td>
                        <td class="DashMOC_header DashMOC_Heading">Current Status</td>
                        <td class="DashMOC_header DashMOC_Heading">Initiation Date</td>
                        <td class="DashMOC_header DashMOC_Heading">Construction Date</td>
                        <td class="DashMOC_header DashMOC_Heading">Start Up Date</td>
                        <td class="DashMOC_header DashMOC_Heading">Submission Date</td>
                        <td class="DashMOC_header DashMOC_Heading">Remarks</td>
                    </tr>
                    <tr runat="server" id="itemPlaceholder">
                    </tr>
                </table>
            </LayoutTemplate>
             <ItemTemplate>
                <tr>
                    <td class="textAlign">
                        <asp:Label ID="lblCAFNo" runat="server" Text='<%# Eval("CafNo") %>' Visible="True"></asp:Label>
                    </td>
                    <td class="textAlign">
                         <asp:Label ID="lblCAFType" runat="server" Text='<%# Eval("CAFType") %>' Visible="True" CssClass="textAlign"></asp:Label>
                    </td>
                    <td class="textAlign">
                        <asp:Label ID="lblCAFTitle" runat="server" Text='<%# Eval("Title") %>' Visible="True" CssClass="textAlign"></asp:Label>
                    </td>
                     <td class="textAlign">
                        <asp:Label ID="lblCRP" runat="server" Text='<%# Eval("CRPName") %>' Visible="True" CssClass="textAlign"></asp:Label>
                    </td>
                     <td class="textAlign">
                        <asp:Label ID="lblCurrentStatus" runat="server" Text='<%# Eval("CurrentSection") %>' Visible="True" CssClass="textAlign"></asp:Label>
                    </td>
                     <td class="textAlign">
                        <asp:Label ID="lblInitiationDate" runat="server" Text='<%# Eval("InitiationDate") %>' Visible="True" CssClass="textAlign"></asp:Label>
                     </td>
                     <td class="textAlign">
                        <asp:Label ID="lblConstructionDate" runat="server" Text='<%# Eval("ConstructionDate") %>' Visible="True" CssClass="textAlign"></asp:Label>
                     </td>
                     <td class="textAlign">
                        <asp:Label ID="lblStartUpDate" runat="server" Text='<%# Eval("StartUpDate") %>' Visible="True" CssClass="textAlign"></asp:Label>
                     </td>
                     <td class="textAlign">
                        <asp:Label ID="Label3" runat="server" Text='<%# Eval("SubmissionDate") %>' Visible="True" CssClass="textAlign"></asp:Label>
                     </td>
                     <td class="textAlign">
                        <asp:Label ID="lblRemarks" runat="server" Text='<%# Eval("Remarks") %>' Visible="True" CssClass="textAlign"></asp:Label>
                     </td>
                </tr>
            </ItemTemplate>
        </asp:ListView>
      </td>
    </tr>
</table>