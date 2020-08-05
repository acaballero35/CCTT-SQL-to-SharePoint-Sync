using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Security;
using CCTT_SQL_to_SharePoint_Sync.ContactTraceDataSetTableAdapters;
using Microsoft.SharePoint.Client;
using static CCTT_SQL_to_SharePoint_Sync.ContactTraceDataSet;

namespace CCTT_SQL_to_SharePoint_Sync
{
    class Program
    {
        static string strLog = DateTime.Now.ToString();
        static string strLogItem;

        public static void Main(string[] args)
        {
            // connect to SharePoint
            using (ClientContext clientContext = new ClientContext("https://cccedu.sharepoint.com/sites/CCTTManagement"))
            {
                try
                {
                    // credentials
                    SecureString securePassword = new SecureString();
                    foreach (char chr in Properties.Settings.Default.SPPwd)
                    {
                        securePassword.AppendChar(chr);
                    }
                    clientContext.Credentials = new SharePointOnlineCredentials("cccautomatedmessaging@cccedu.onmicrosoft.com", securePassword);

                    processCompany(clientContext);
                    processRegistrations(clientContext);
                }
                catch (Exception ex)
                {
                    strLogItem += "Error: " + ex.Message;
                }

                strLog += "-" + DateTime.Now.ToString();

                // get log list
                List lList = clientContext.Web.Lists.GetByTitle("Sync Log SQL to SP");
                clientContext.Load(lList);
                clientContext.ExecuteQuery();

                ListItem litem = lList.AddItem(new ListItemCreationInformation());
                litem["Title"] = strLog;
                litem["Log"] = strLogItem;
                litem.Update();
                clientContext.Load(litem);
                clientContext.ExecuteQuery();
            }
        }

        static void processCompany(ClientContext clientContext)
        {
            // get Company data from sql db into data table
            companyTableAdapter cda = new companyTableAdapter();
            companyDataTable cdt = new companyDataTable();
            _ = cda.Fill(cdt);

            strLogItem += "Company table rows read: " + cdt.Rows.Count + "\r\n";
            int rowsProcessed = 0;

            // get list to sync
            List oList = clientContext.Web.Lists.GetByTitle("Company");
            CamlQuery cq = new CamlQuery();
            cq.ViewXml = "<View/>";

            // get list item collection
            ListItemCollection items = oList.GetItems(cq);
            clientContext.Load(items);
            clientContext.ExecuteQuery();

            // cast list item collection to dictionary for fast finding
            Dictionary<string, ListItem> CompanyItems = items.Cast<ListItem>().ToDictionary(i => (string)i["CompanyID"], i => i);

            // loop through db rows
            foreach (DataRow row in cdt.Rows)
            {
                try
                {
                    // check if the row is already a list item in sp list
                    if (CompanyItems.TryGetValue(row["Id"].ToString(), out ListItem item))
                    {
                        // if there are changes in sp list item then apply to row
                        if (
                            row["name"].ToString() != item["Title"].ToString() ||
                            row["contactEmail"].ToString() != item["contactEmail"].ToString() ||
                            row["contactName"].ToString() != item["contactName"].ToString()
                            )
                        {
                            row["name"] = item["Title"].ToString();
                            row["contactEmail"] = item["contactEmail"].ToString();
                            row["contactName"] = item["contactName"].ToString();
                        }
                    }
                    // if the row is not a list item in sp list, add it
                    else
                    {
                        ListItem nitem = oList.AddItem(new ListItemCreationInformation());

                        nitem["CompanyID"] = row["Id"].ToString();
                        nitem["Title"] = row["name"].ToString();
                        nitem["contactEmail"] = row["contactEmail"].ToString();
                        nitem["contactName"] = row["contactName"].ToString();
                        nitem["copiedfromdb"] = DateTime.Now.AddHours(-5);

                        nitem.Update();
                        clientContext.Load(nitem);
                        clientContext.ExecuteQuery();
                    }

                    rowsProcessed++;
                }
                catch (Exception ex)
                {
                    strLogItem += "Error: " + ex.Message + "\r\n";
                }
            }

            strLogItem += "Company table rows processed: " + rowsProcessed + "\r\n";
        }

        static void processRegistrations(ClientContext clientContext)
        {
            // get Registrations data from sql db into data table
            registrationsTableAdapter rda = new registrationsTableAdapter();
            registrationsDataTable rdt = new registrationsDataTable();

            _ = rda.Fill(rdt);

            strLogItem += "Registrations table rows read: " + rdt.Rows.Count + "\r\n";
            int rowsProcessed = 0;

            // get list to sync
            List oList = clientContext.Web.Lists.GetByTitle("Registrations");
            CamlQuery cq = new CamlQuery();
            cq.ViewXml = "<View/>";

            // get list item collection
            ListItemCollection items = oList.GetItems(cq);
            clientContext.Load(items);
            clientContext.ExecuteQuery();

            // cast list item collection to dictionary for fast finding
            Dictionary<string, ListItem> RegistrationItems = items.Cast<ListItem>().ToDictionary(i => (string)i["RegistrationID"], i => i);

            // get Company list to find SP lookup ID
            List oListCo = clientContext.Web.Lists.GetByTitle("Company");
            CamlQuery cqCo = new CamlQuery();
            cqCo.ViewXml = "<View/>";

            // get Company list item collection to find SP lookup ID
            ListItemCollection itemsCo = oListCo.GetItems(cqCo);
            clientContext.Load(itemsCo);
            clientContext.ExecuteQuery();

            // cast Company list item collection to dictionary for fast finding of SP lookup ID
            Dictionary<string, ListItem> CompanyItems = itemsCo.Cast<ListItem>().ToDictionary(i => (string)i["CompanyID"], i => i);

            // loop through db rows
            foreach (DataRow row in rdt.Rows)
            {
                try
                {
                    // check if the row is already a list item in sp list
                    if (RegistrationItems.TryGetValue(row["id"].ToString(), out ListItem item))
                    {
                        // if there are changes in sp list item then apply to row
                        if (
                            row["personalEmail"].ToString() != item["Title"].ToString() ||
                            row["lastName"].ToString() != item["lastName"].ToString() ||
                            row["firstName"].ToString() != item["firstName"].ToString() ||
                            row["MI"].ToString() != (item["MI"] == null ? "" : item["MI"].ToString()) ||
                            row["birthDate"].ToString() != item["birthDate"].ToString() ||
                            row["gender"].ToString() != item["gender"].ToString() ||
                            row["address1"].ToString() != item["address1"].ToString() ||
                            row["address2"].ToString() != (item["address2"] == null ? "" : item["address2"].ToString()) ||
                            row["city"].ToString() != item["city"].ToString() ||
                            row["state"].ToString() != item["state"].ToString() ||
                            row["zip"].ToString() != item["zip"].ToString() ||
                            row["suffix"].ToString() != (item["suffix"] == null ? "" : item["suffix"].ToString()) ||
                            row["primaryPhone"].ToString() != item["primaryPhone"].ToString() ||
                            row["phoneType"].ToString() != item["phoneType"].ToString() ||
                            row["language"].ToString() != item["language"].ToString() ||
                            row["studentSignature"].ToString() != (item["studentSignature"] == null ? "" : item["studentSignature"].ToString()) ||
                            row["registeredDate"].ToString() != item["registeredDate"].ToString()
                            )
                        {
                            row["Title"] = item["personalEmail"].ToString();
                            row["lastName"] = item["lastName"].ToString();
                            row["firstName"] = item["firstName"].ToString();
                            row["MI"] = item["MI"] == null ? "" : item["MI"].ToString();
                            row["birthDate"] = item["birthDate"].ToString();
                            row["gender"] = item["gender"].ToString();
                            row["address1"] = item["address1"].ToString();
                            row["address2"] = item["address2"] == null ? "" : item["address2"].ToString();
                            row["city"] = item["city"].ToString();
                            row["state"] = item["state"].ToString();
                            row["zip"] = item["zip"].ToString();
                            row["suffix"] = item["suffix"].ToString();
                            row["primaryPhone"] = item["primaryPhone"].ToString();
                            row["phoneType"] = item["phoneType"].ToString();
                            row["language"] = item["language"].ToString();
                            row["studentSignature"] = item["studentSignature"] == null ? "" : item["studentSignature"].ToString();
                            row["registeredDate"] = item["registeredDate"].ToString();
                        }
                    }
                    // if the row is not a list item in sp list, add it
                    else
                    {
                        CompanyItems.TryGetValue(row["employer"].ToString(), out ListItem itemCo);

                        ListItem nitem = oList.AddItem(new ListItemCreationInformation());

                        nitem["employer"] = itemCo["ID"];
                        nitem["Title"] = row["personalEmail"].ToString();
                        nitem["RegistrationID"] = row["id"].ToString();
                        nitem["lastName"] = row["lastName"].ToString();
                        nitem["firstName"] = row["firstName"].ToString();
                        nitem["MI"] = row["MI"].ToString();
                        nitem["birthDate"] = row["birthDate"].ToString();
                        nitem["gender"] = row["gender"].ToString();
                        nitem["address1"] = row["address1"].ToString();
                        nitem["address2"] = row["address2"].ToString();
                        nitem["city"] = row["city"].ToString();
                        nitem["state"] = row["state"].ToString();
                        nitem["zip"] = row["zip"].ToString();
                        nitem["suffix"] = row["suffix"].ToString();
                        nitem["primaryPhone"] = row["primaryPhone"].ToString();
                        nitem["phoneType"] = row["phoneType"].ToString();
                        nitem["language"] = row["language"].ToString();
                        nitem["studentSignature"] = row["studentSignature"].ToString();
                        nitem["registeredDate"] = row["registeredDate"].ToString();
                        nitem["copiedfromdb"] =  DateTime.Now.AddHours(-5);

                        nitem.Update();
                        clientContext.Load(nitem);
                        clientContext.ExecuteQuery();
                    }

                    rowsProcessed++;
                }
                catch (Exception ex)
                {
                    strLogItem += "Error: " + ex.Message + "\r\n";
                }

            }

            strLogItem += "Registrations table rows processed: " + rowsProcessed + "\r\n";
        }
    }
}
