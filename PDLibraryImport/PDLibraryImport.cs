using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.DocumentManagement.DocumentSets;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Security;
using Microsoft.SharePoint.Utilities;


namespace PDLibraryImport
{
    public partial class PDLibraryImport : Form
    {
        public PDLibraryImport()
        {
            InitializeComponent();
        }

        public static string StringConnection()
        {
            return "Server=Vdbw001tv\\msapps;Database=Staging;Trusted_Connection=True;";
        }

        public static DataSet getQuery(string query)
        {
            using (SqlConnection dbConnection = new SqlConnection(StringConnection()))
            {

                dbConnection.Open();

                SqlDataAdapter objCmd = new SqlDataAdapter(query, dbConnection);
                DataSet objDS = new DataSet();
                objCmd.Fill(objDS, "Data");

                dbConnection.Close();

                return objDS;

            }

        }

        private void btnStart_Click(object sender, EventArgs e)
        {

            string url = "http://spmain.volpe.dot.gov";
            //string url = "http://spmaindev.volpe.dot.gov";
            //string url = "http://zebaduag03644";

            //string sub = "*";
            string sub = "/sites/Tools/HumanResources/p21/PDLibrary";

            using (SPSite site = new SPSite(url + sub))
            {
                //using (SPWeb web = (sub == "*") ? site.RootWeb : site.OpenWeb(sub))
                string justForWarmup = site.Url;

                //using (SPWeb web = site.OpenWeb(sub))
                using (SPWeb web = site.OpenWeb())
                {
                    try
                    {
                        justForWarmup = web.Url;

                        //Open library
                        string fullLibraryUrl = web.Url + "/Documents/";
                        SPDocumentLibrary list = (SPDocumentLibrary)web.GetList(fullLibraryUrl);

                        //Find the content type to use
                        SPContentType docsetCT = list.ContentTypes["Document Set"];

                        //get data from database
                        string PDQuery =
                        "SELECT [PDF].* " +
                        "FROM " +
                        "       (" +
                        "       SELECT [PD].[id], [PD].[Create_Date], [PD].[Position_Number], [PDT].[Title], [PD].[Organization_Id], [PD].[Pay_Grade], NULL AS 'Supervisory_Position', [PD].[Notes], [PD].[Status], [PD].[Update_Date] " +
                        "       FROM   [Staging].[dbo].[Position_Description] [PD] WITH (NOLOCK) " +
                        "       JOIN   [Staging].[dbo].[Position_Description_Document_JOIN] [PDDJ] WITH (NOLOCK) ON [PD].[id] = [PDDJ].[Position_Description_Id] AND [PDDJ].[File_Type] = 'pdf' " +
                        "       JOIN   [Staging].[dbo].[Position_Description_Title] [PDT] ON [PD].Title_Id = [PDT].[id] " +
                        "       WHERE  LEFT([Position_Number], 1) IN ('A', 'P', 'T') " +
                        "       ) AS [PDF] " +
                        "JOIN " +
                        "       (" +
                        "       SELECT [PD].[id], [PD].[Create_Date], [PD].[Position_Number], [PDT].[Title], [PD].[Organization_Id], [PD].[Pay_Grade], NULL AS 'Supervisory_Position', [PD].[Notes], [PD].[Status], [PD].[Update_Date] " +
                        "       FROM   [Staging].[dbo].[Position_Description] [PD] WITH (NOLOCK) " +
                        "       JOIN   [Staging].[dbo].[Position_Description_Document_JOIN] [PDDJ] WITH (NOLOCK) ON  [PD].[id] = [PDDJ].[Position_Description_Id] AND [PDDJ].[File_Type] = 'doc' " +
                        "       JOIN   [Staging].[dbo].[Position_Description_Title] [PDT] WITH (NOLOCK) ON [PD].Title_Id = [PDT].[id] " +
                        "       WHERE  LEFT([Position_Number], 1) IN ('A', 'P', 'T') " +
                        "       ) AS [DOC] " +
                        "ON     [PDF].[id] = [DOC].[id] " +
                        "UNION " +
                        "SELECT [PDF].* " +
                        "FROM " +
                        "       (" +
                        "       SELECT [PD].[id], [PD].[Create_Date], [PD].[Position_Number], [PDT].[Title], [PD].[Organization_Id], [PD].[Pay_Grade], NULL AS 'Supervisory_Position', [PD].[Notes], [PD].[Status], [PD].[Update_Date] " +
                        "       FROM   [Staging].[dbo].[Position_Description] [PD] WITH (NOLOCK) " +
                        "       JOIN   [Staging].[dbo].[Position_Description_Document_JOIN] [PDDJ] WITH (NOLOCK) ON [PD].[id] = [PDDJ].[Position_Description_Id] AND [PDDJ].[File_Type] = 'pdf' " +
                        "       JOIN   [Staging].[dbo].[Position_Description_Title] [PDT] WITH (NOLOCK) ON [PD].Title_Id = [PDT].[id] " +
                        "       WHERE  LEFT([Position_Number], 1) IN ('A', 'P', 'T') " +
                        "       ) AS [PDF] " +
                        "JOIN " +
                        "       (" +
                        "       SELECT [PD].[id], [PD].[Create_Date], [PD].[Position_Number], [PDT].[Title], [PD].[Organization_Id], [PD].[Pay_Grade], NULL AS 'Supervisory_Position', [PD].[Notes], [PD].[Status], [PD].[Update_Date] " +
                        "       FROM   [Staging].[dbo].[Position_Description] [PD] WITH (NOLOCK) " +
                        "       JOIN   [Staging].[dbo].[Position_Description_Document_JOIN] [PDDJ] WITH (NOLOCK) ON [PD].[id] = [PDDJ].[Position_Description_Id] AND [PDDJ].[File_Type] = 'docx' " +
                        "       JOIN   [Staging].[dbo].[Position_Description_Title] [PDT] WITH (NOLOCK) ON [PD].Title_Id = [PDT].[id] " +
                        "       WHERE  LEFT([Position_Number], 1) IN ('A', 'P', 'T') " +
                        "       ) AS [DOCX] " +
                        "ON     [PDF].[id] = [DOCX].[id] " +
                        "UNION " +
                        "SELECT [PDF].* " +
                        "FROM " +
                        "       (" +
                        "       SELECT [PD].[id], [PD].[Create_Date], [PD].[Position_Number], [PDT].[Title], [PD].[Organization_Id], [PD].[Pay_Grade], NULL AS 'Supervisory_Position', [PD].[Notes], [PD].[Status], [PD].[Update_Date] " +
                        "       FROM   [Staging].[dbo].[Position_Description] [PD] WITH (NOLOCK) " +
                        "       JOIN   [Staging].[dbo].[Position_Description_Document_JOIN] [PDDJ] WITH (NOLOCK) ON [PD].[id] = [PDDJ].[Position_Description_Id] AND [PDDJ].[File_Type] = 'pdf' " +
                        "       JOIN   [Staging].[dbo].[Position_Description_Title] [PDT] WITH (NOLOCK) ON [PD].Title_Id = [PDT].[id] " +
                        "	   WHERE  LEFT([Position_Number], 1) IN ('A', 'P', 'T') " +
                        "       ) AS [PDF] " +
                        "JOIN " +
                        "       (" +
                        "       SELECT [PD].[id], [PD].[Create_Date], [PD].[Position_Number], [PDT].[Title], [PD].[Organization_Id], [PD].[Pay_Grade], NULL AS 'Supervisory_Position', [PD].[Notes], [PD].[Status], [PD].[Update_Date] " +
                        "       FROM   [Staging].[dbo].[Position_Description] [PD] WITH (NOLOCK) " +
                        "       JOIN   [Staging].[dbo].[Position_Description_Document_JOIN] [PDDJ] WITH (NOLOCK) ON [PD].[id] = [PDDJ].[Position_Description_Id] AND [PDDJ].[File_Type] = 'rtf' " +
                        "       JOIN   [Staging].[dbo].[Position_Description_Title] [PDT] WITH (NOLOCK) ON [PD].Title_Id = [PDT].[id] " +
                        "       WHERE  LEFT([Position_Number], 1) IN ('A', 'P', 'T') " +
                        "       ) AS [RTF] " +
                        "ON     [PDF].[id] = [RTF].[id] ";

                        DataSet PDData = getQuery(PDQuery);

                        DataTable dataTable = new DataTable();
                        dataTable = PDData.Tables[0];

                        DocumentSet ds = null;

                        foreach (DataRow dataRow in dataTable.Rows)
                        {

                            //Collect documentset properties
                            System.Collections.Hashtable properties = new System.Collections.Hashtable();

                            //Use columns' internal names
                            string docSetName = dataRow["Position_Number"].ToString();

                            properties.Add("PDDateClassified", dataRow["Create_Date"].ToString());
                            properties.Add("PDNumber", dataRow["Position_Number"].ToString());
                            properties.Add("PDTitle", dataRow["Title"].ToString());
                            properties.Add("Title", dataRow["Title"].ToString());

                            SPList orgList = web.Lists["Organizations"];
                            SPQuery query = new SPQuery();
                            int intSelectedId = 0;
                            query.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + "V-" + dataRow["Organization_Id"].ToString() + "</Value></Eq></Where>";
                            if (orgList.GetItems(query) != null)
                            {
                                try
                                {
                                    SPListItem result = orgList.GetItems(query)[0];
                                    intSelectedId = result.ID;
                                }
                                catch (Exception e1)
                                {
                                    intSelectedId = 0;
                                }
                            }
                            else
                            {
                                intSelectedId = 0;
                            }

                            if (intSelectedId > 0)
                            {
                                properties.Add("PDOrg", new SPFieldLookupValue(intSelectedId, "V-" + dataRow["Organization_Id"].ToString()));
                            }
                            else
                            {
                                properties.Add("PDOrg", new SPFieldLookupValue(1, "V-100"));
                            }

                            properties.Add("PDSeries", new SPFieldLookupValue(1, "0018")); //I chose this static "default" value since there is no legacy data for this attribute

                            SPList payGradesList = web.Lists["PayPlansGrades"];
                            SPQuery query2 = new SPQuery();
                            int intSelectedId2 = 0;
                            query2.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='TEXT'>" + dataRow["Pay_Grade"].ToString() + "</Value></Eq></Where>";
                            if (payGradesList.GetItems(query2) != null)
                            {
                                try
                                {
                                    SPListItem result = payGradesList.GetItems(query2)[0];
                                    intSelectedId2 = result.ID;
                                }
                                catch (Exception e2)
                                {
                                    intSelectedId2 = 0;
                                }
                            }
                            else
                            {
                                intSelectedId2 = 0;
                            }

                            if (intSelectedId2 > 0)
                            {
                                properties.Add("PDPayGrade", new SPFieldLookupValue(intSelectedId2, dataRow["Pay_Grade"].ToString()));
                            }
                            else
                            {
                                properties.Add("PDPayGrade", new SPFieldLookupValue(1, "ED-00"));
                            }

                            properties.Add("PDSupervisoryPosition", "No"); //I chose this static "default" value since there is no legacy data for this attribute
                            properties.Add("PDNotes", dataRow["Notes"].ToString());
                            properties.Add("PDVisibility", dataRow["Status"].ToString());
                            properties.Add("PDDateUploaded", dataRow["Update_Date"].ToString());

                            //Create documentset if it doesn't exist already
                            if (!list.ParentWeb.GetFolder(SPUrlUtility.CombineUrl(list.RootFolder.ServerRelativeUrl, docSetName)).Exists)
                            {
                                ds = DocumentSet.Create(list.RootFolder, docSetName, docsetCT.Id, properties, true);
                                rtb.AppendText("Created document set " + docSetName + "\r\n");
                            }

                            //Get files to add to documentset
                            string PDDocumentsQuery =
                            "SELECT [FL].[Position_Number], [PDDJ].[Short_File_Name], [PDDJ].[Document_Type], [PDDJ].[Archived_Flag], [PDD].[Content] " +
                            "FROM " +
                            "(SELECT [PDF].* " +
                            "FROM " +
                            "       (" +
                            "       SELECT [PD].[id], [PD].[Create_Date], [PD].[Position_Number], [PDT].[Title], [PD].[Organization_Id], [PD].[Pay_Grade], NULL AS 'Supervisory_Position', [PD].[Notes], [PD].[Status], [PD].[Update_Date] " +
                            "       FROM   [Staging].[dbo].[Position_Description] [PD] WITH (NOLOCK) " +
                            "       JOIN   [Staging].[dbo].[Position_Description_Document_JOIN] [PDDJ] WITH (NOLOCK) ON [PD].[id] = [PDDJ].[Position_Description_Id] AND [PDDJ].[File_Type] = 'pdf' " +
                            "	   JOIN   [Staging].[dbo].[Position_Description_Title] [PDT] ON [PD].Title_Id = [PDT].[id] " +
                            "       WHERE  LEFT([Position_Number], 1) IN ('A', 'P', 'T') " +
                            "       ) AS [PDF] " +
                            "JOIN " +
                            "       (" +
                            "       SELECT [PD].[id], [PD].[Create_Date], [PD].[Position_Number], [PDT].[Title], [PD].[Organization_Id], [PD].[Pay_Grade], NULL AS 'Supervisory_Position', [PD].[Notes], [PD].[Status], [PD].[Update_Date] " +
                            "       FROM   [Staging].[dbo].[Position_Description] [PD] WITH (NOLOCK) " +
                            "       JOIN   [Staging].[dbo].[Position_Description_Document_JOIN] [PDDJ] WITH (NOLOCK) ON  [PD].[id] = [PDDJ].[Position_Description_Id] AND [PDDJ].[File_Type] = 'doc' " +
                            "	   JOIN   [Staging].[dbo].[Position_Description_Title] [PDT] WITH (NOLOCK) ON [PD].Title_Id = [PDT].[id] " +
                            "       WHERE  LEFT([Position_Number], 1) IN ('A', 'P', 'T') " +
                            "       ) AS [DOC] " +
                            "ON     [PDF].[id] = [DOC].[id] " +
                            "UNION " +
                            "SELECT [PDF].* " +
                            "FROM " +
                            "       (" +
                            "       SELECT [PD].[id], [PD].[Create_Date], [PD].[Position_Number], [PDT].[Title], [PD].[Organization_Id], [PD].[Pay_Grade], NULL AS 'Supervisory_Position', [PD].[Notes], [PD].[Status], [PD].[Update_Date] " +
                            "       FROM   [Staging].[dbo].[Position_Description] [PD] WITH (NOLOCK) " +
                            "       JOIN   [Staging].[dbo].[Position_Description_Document_JOIN] [PDDJ] WITH (NOLOCK) ON [PD].[id] = [PDDJ].[Position_Description_Id] AND [PDDJ].[File_Type] = 'pdf' " +
                            "	   JOIN   [Staging].[dbo].[Position_Description_Title] [PDT] WITH (NOLOCK) ON [PD].Title_Id = [PDT].[id] " +
                            "       WHERE  LEFT([Position_Number], 1) IN ('A', 'P', 'T') " +
                            "       ) AS [PDF] " +
                            "JOIN " +
                            "       (" +
                            "       SELECT [PD].[id], [PD].[Create_Date], [PD].[Position_Number], [PDT].[Title], [PD].[Organization_Id], [PD].[Pay_Grade], NULL AS 'Supervisory_Position', [PD].[Notes], [PD].[Status], [PD].[Update_Date] " +
                            "       FROM   [Staging].[dbo].[Position_Description] [PD] WITH (NOLOCK) " +
                            "       JOIN   [Staging].[dbo].[Position_Description_Document_JOIN] [PDDJ] WITH (NOLOCK) ON [PD].[id] = [PDDJ].[Position_Description_Id] AND [PDDJ].[File_Type] = 'docx' " +
                            "	   JOIN   [Staging].[dbo].[Position_Description_Title] [PDT] WITH (NOLOCK) ON [PD].Title_Id = [PDT].[id] " +
                            "       WHERE  LEFT([Position_Number], 1) IN ('A', 'P', 'T') " +
                            "       ) AS [DOCX] " +
                            "ON     [PDF].[id] = [DOCX].[id] " +
                            "UNION " +
                            "SELECT [PDF].* " +
                            "FROM " +
                            "       (" +
                            "       SELECT [PD].[id], [PD].[Create_Date], [PD].[Position_Number], [PDT].[Title], [PD].[Organization_Id], [PD].[Pay_Grade], NULL AS 'Supervisory_Position', [PD].[Notes], [PD].[Status], [PD].[Update_Date] " +
                            "       FROM   [Staging].[dbo].[Position_Description] [PD] WITH (NOLOCK) " +
                            "       JOIN   [Staging].[dbo].[Position_Description_Document_JOIN] [PDDJ] WITH (NOLOCK) ON [PD].[id] = [PDDJ].[Position_Description_Id] AND [PDDJ].[File_Type] = 'pdf' " +
                            "       JOIN   [Staging].[dbo].[Position_Description_Title] [PDT] WITH (NOLOCK) ON [PD].Title_Id = [PDT].[id] " +
                            "	   WHERE  LEFT([Position_Number], 1) IN ('A', 'P', 'T') " +
                            "       ) AS [PDF] " +
                            "JOIN " +
                            "       (" +
                            "       SELECT [PD].[id], [PD].[Create_Date], [PD].[Position_Number], [PDT].[Title], [PD].[Organization_Id], [PD].[Pay_Grade], NULL AS 'Supervisory_Position', [PD].[Notes], [PD].[Status], [PD].[Update_Date] " +
                            "       FROM   [Staging].[dbo].[Position_Description] [PD] WITH (NOLOCK)  " +
                            "       JOIN   [Staging].[dbo].[Position_Description_Document_JOIN] [PDDJ] WITH (NOLOCK) ON [PD].[id] = [PDDJ].[Position_Description_Id] AND [PDDJ].[File_Type] = 'rtf' " +
                            "	   JOIN   [Staging].[dbo].[Position_Description_Title] [PDT] WITH (NOLOCK) ON [PD].Title_Id = [PDT].[id] " +
                            "       WHERE  LEFT([Position_Number], 1) IN ('A', 'P', 'T') " +
                            "       ) AS [RTF] " +
                            "ON     [PDF].[id] = [RTF].[id]) AS [FL] " +
                            "JOIN [Staging].[dbo].[Position_Description_Document_JOIN] [PDDJ] ON [FL].[id] = [PDDJ].[Position_Description_Id] " +
                            "JOIN [Staging].[dbo].[Position_Description_Document] [PDD] ON [PDD].[id] = [PDDJ].[Position_Description_Document_Id] " +
                            "WHERE [FL].[Position_Number] = '" + dataRow["Position_Number"].ToString() + "' ";
                            //"AND [PDDJ].Archived_Flag = 'N' ";

                            DataSet PDDocuments = getQuery(PDDocumentsQuery);

                            DataTable documentsDataTable = new DataTable();
                            documentsDataTable = PDDocuments.Tables[0];

                            //Add files to documentset
                            SPFolder docSet = list.ParentWeb.GetFolder(SPUrlUtility.CombineUrl(list.RootFolder.ServerRelativeUrl, docSetName));

                            foreach (DataRow docsRow in documentsDataTable.Rows)
                            {
                                System.Collections.Hashtable properties2 = new System.Collections.Hashtable();
                                properties2.Add("PDDocumentType", docsRow["Document_Type"].ToString());

                                SPFile pdFile = docSet.Files.Add(docsRow["Short_File_Name"].ToString(), (byte[])docsRow["Content"], properties2, true);
                                rtb.AppendText("Added file " + docsRow["Short_File_Name"].ToString() + " to document set " + docSetName + "\r\n");
                            }

                            docSetName = "";
                            ds = null;
                        }

                        rtb.AppendText("\r\n" + "Export done. Refreshing..." + "\r\n");
                        list.Update();

                    }
                    catch (Exception ex)
                    {
                        SPDiagnosticsService diagSvc = SPDiagnosticsService.Local;
                        diagSvc.WriteTrace(0, new SPDiagnosticsCategory("PDLibrary", TraceSeverity.Monitorable, EventSeverity.Error),
                        TraceSeverity.Monitorable, "PD Library error:  {0}", new object[] { ex.ToString() });
                        rtb.AppendText(ex.ToString() + "\r\n");
                    }

                }
            }

        }
    }
}


