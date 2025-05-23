using Logger;
using System;
using System.Collections.Generic;
using System.Data;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomainController
{
    public class Group
    {
        public static string logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\AssertYIT\\DomainController\\";
        LoggerClass objLoggerClass = new LoggerClass("");
        LogWriter objLogWriter = new LogWriter(logFilePath);
        helperMethod objHelper = new helperMethod();
        bool isConnect = false;
        DirectoryEntry entry;
        DirectorySearcher srch;
        string DomainName = "";
        string DomainUsername = "";
        string DomainUserPassword = "";
        public Group(string DomainName, string Username, string Password)
        {
            this.DomainName = DomainName;
            this.DomainUsername = Username;
            this.DomainUserPassword = Password;
        }
        public bool Connect()
        {
            try
            {
                entry = new DirectoryEntry(string.Format("LDAP://{0}", DomainName), DomainUsername, DomainUserPassword);
                srch = new DirectorySearcher(entry);
                isConnect = true;
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("Group", "Connect", 0, "can not connect to server", "exception:" + ex);objLogWriter.WriteLogsFromQueue();
            }
            return isConnect;
        }
        public bool createGroup(string grpName, string grpDescription)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        if (objHelper.isGrpExists(entry, grpName))
                        {
                            using (DirectoryEntry grpEntry = entry.Children.Add("CN=iGroup", "group"))
                            {
                                grpEntry.Properties["sAMAccountName"].Value = grpName;
                                grpEntry.Properties["description"].Value = grpDescription;
                                grpEntry.CommitChanges();
                                bl = true;
                            }
                        }
                        else
                        {
                            objLoggerClass.Write("CreateGroup", "GroupExists", 0, $"Group '{grpName}' already exists", "Group creation failed due to duplicate"); objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("CreateGroup", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("CreateGroup", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("createGrp", "createGroup", 0, "Create group in DC", "Exception : " + ex);
                objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }


        public bool deleteGroup(string grpName)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        srch.Filter = $"(sAMAccountName={grpName})";
                        SearchResult result = srch.FindOne();
                        if (result != null)
                        {
                            DirectoryEntry userEntry = result.GetDirectoryEntry();
                            userEntry.DeleteTree();
                            userEntry.CommitChanges();
                            bl = true;
                        }
                        else
                        {
                            objLoggerClass.Write("deletegrp", "notFound", 0, "delete group in DC", "user not found");objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("deletegrp", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("deletegrp", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                } 
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("deletegrp", "deleteGroup", 0, " delete group in DC", "Exception : " + ex);
                objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }


        public DataTable grpInformation(string grpName)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("GroupName");
            dt.Columns.Add("Description");
            dt.Columns.Add("DistinguishedName");
            dt.Columns.Add("GroupType");
            dt.Columns.Add("Members");

            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        srch.Filter = $"(&(objectClass=group)(sAMAccountName={grpName}))";
                        srch.SearchScope = SearchScope.Subtree;

                        SearchResult result = srch.FindOne();
                        if (result != null)
                        {
                            DirectoryEntry grpEntry = result.GetDirectoryEntry();
                            DataRow row = dt.NewRow();
                            row["GroupName"] = grpEntry.Properties["sAMAccountName"].Value;
                            row["Description"] = grpEntry.Properties["description"].Value;
                            row["DistinguishedName"] = grpEntry.Properties["distinguishedName"].Value;
                            row["GroupType"] = grpEntry.Properties["groupType"].Value;

                            string members = string.Empty;
                            foreach (var member in grpEntry.Properties["member"])
                            {
                                members += member + "; ";
                            }
                            row["Members"] = members;
                            dt.Rows.Add(row);
                        }
                        else
                        {
                            objLoggerClass.Write("groupInfo", "grpInformation", 0, " get information of group in DC", "group not found");
                            objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("groupInfo", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("groupInfo", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("groupInfo", "grpInformation", 0, " get information of group in DC", "Exception : " + ex);
                objLogWriter.WriteLogsFromQueue();
                return null;
            }
            return dt;
        }


        public bool addUserToGrp(string userName, string grpName)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                            DirectoryEntry grpEntry = null;
                            using (DirectorySearcher groupSearch = new DirectorySearcher(entry))
                            {
                                groupSearch.Filter = $"(&(objectClass=group)(sAMAccountName={grpName}))";
                                groupSearch.SearchScope = SearchScope.Subtree;

                                SearchResult grpResult = groupSearch.FindOne();
                                if (grpResult != null)
                                {
                                    grpEntry = grpResult.GetDirectoryEntry();
                                }
                                else
                                {
                                    objLoggerClass.Write("addUser", "addUserToGrp", 0, " add user in group in DC", "group not found");
                                    objLogWriter.WriteLogsFromQueue();
                                }
                            }
                            DirectoryEntry userEntry = null;
                            using (DirectorySearcher userSearch = new DirectorySearcher(entry))
                            {
                                userSearch.Filter = $"(&(objectClass=user)(sAMAccountName={userName}))";
                                userSearch.SearchScope = SearchScope.Subtree;

                                SearchResult userResult = userSearch.FindOne();
                                if (userResult != null)
                                {
                                    userEntry = userResult.GetDirectoryEntry();
                                }
                                else
                                {
                                    objLoggerClass.Write("addUser", "addUserToGrp", 0, " add user in group in DC", "user not found");
                                    objLogWriter.WriteLogsFromQueue();
                                }
                            }
                            grpEntry.Properties["member"].Add(userEntry.Properties["distinguishedName"].Value);
                            grpEntry.CommitChanges();
                            bl= true;
                    }
                    else
                    {
                        objLoggerClass.Write("addUser", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("addUser", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("addUser", "addUserToGrp", 0, "add user in group in DC", "Exception : " + ex); objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }


        public bool dltUserFromGrp(string userName, string grpName)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                            DirectoryEntry grpEntry = null;
                            using (DirectorySearcher groupSearch = new DirectorySearcher(entry))
                            {
                                groupSearch.Filter = $"(&(objectClass=group)(sAMAccountName={grpName}))";
                                groupSearch.SearchScope = SearchScope.Subtree;

                                SearchResult grpResult = groupSearch.FindOne();
                                if (grpResult != null)
                                {
                                    grpEntry = grpResult.GetDirectoryEntry();
                                }
                                else
                                {
                                    objLoggerClass.Write("dltUser", "dltUserFromGrp", 0, "delete user from group in DC", "group not found");
                                    objLogWriter.WriteLogsFromQueue();
                                }
                            }
                            DirectoryEntry userEntry = null;
                            using (DirectorySearcher userSearch = new DirectorySearcher(entry))
                            {
                                userSearch.Filter = $"(&(objectClass=user)(sAMAccountName={userName}))";
                                userSearch.SearchScope = SearchScope.Subtree;

                                SearchResult userResult = userSearch.FindOne();
                                if (userResult != null)
                                {
                                    userEntry = userResult.GetDirectoryEntry();
                                }
                                else
                                {
                                    objLoggerClass.Write("dltUser", "dltUserFromGrp", 0, "delete user from group in DC", "user not found");
                                    objLogWriter.WriteLogsFromQueue();
                                }
                            }
                            grpEntry.Properties["member"].Remove(userEntry.Properties["distinguishedName"].Value);
                            grpEntry.CommitChanges();
                            bl = true;
                        
                    }
                    else 
                    {
                        objLoggerClass.Write("dltUser", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else 
                {
                    objLoggerClass.Write("dltUser", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("dltUser", "dltUserFromGrp", 0, "delete user from group in DC", "Exception : " + ex);objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }


        public bool moveGrpToOU(string grpName, string targetOUDN)
        {
            bool bl = false;
            try
            {
                using (DirectoryEntry rootDSE = new DirectoryEntry($"LDAP://{DomainName}/rootDSE", DomainUsername, DomainUserPassword))
                {
                    string defaultNamingContext = rootDSE.Properties["defaultNamingContext"].Value.ToString();
                    if (Connect())
                    {
                        if (objHelper.AreCredentialValid(entry))
                        {
                            srch.Filter = $"(sAMAccountName={grpName})";
                            SearchResult result = srch.FindOne();
                            if (result != null)
                            {
                                using (DirectoryEntry userEntry = result.GetDirectoryEntry())
                                {
                                    using (DirectoryEntry targetOU = new DirectoryEntry($"LDAP://{DomainName}/{targetOUDN},{defaultNamingContext}", DomainUsername, DomainUserPassword))
                                    {
                                        if(targetOU != null && targetOU.NativeObject != null)
                                        {
                                            userEntry.MoveTo(targetOU);
                                            userEntry.CommitChanges();
                                            return true;
                                        }
                                        else
                                        {
                                            objLoggerClass.Write("moveGroup", "notFound", 0, "move group in OU in DC", "targetOU not found");objLogWriter.WriteLogsFromQueue();
                                        }
                                    }
                                }
                            }
                            else
                            {
                                objLoggerClass.Write("moveGroup", "notFound", 0, "move group in OU in DC", "group not found");objLogWriter.WriteLogsFromQueue();
                            }
                        }
                        else
                        {
                            objLoggerClass.Write("moveGroup", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("moveGroup", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                    }
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("moveGroup", "moveGrpToOU", 0, "move group in OU in DC", "Exception : " + ex);
                objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }

    }
}
