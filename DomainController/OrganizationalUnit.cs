using Logger;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomainController
{
    public class OrganizationalUnit
    {
        public static string logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\AssertYIT\\DomainController\\";
        LoggerClass objLoggerClass = new LoggerClass("");
        LogWriter objLogWriter = new LogWriter(logFilePath);
        bool isConnect = false;
        helperMethod objHelper = new helperMethod();
        DirectoryEntry entry;
        DirectorySearcher srch;

        string DomainName = "";
        string DomainUsername = "";
        string DomainUserPassword = "";

        public OrganizationalUnit(string DomainName, string Username, string Password)
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
                objLoggerClass.Write("OrganizationalUnit", "Connect", 0, "can not connect to server", "exception:" + ex); objLogWriter.WriteLogsFromQueue();
            }
            return isConnect;
        }

        public bool createOU(string OUName, string OUDescription)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        if(objHelper.IsOUExists(entry, OUName))
                        {
                            using (DirectoryEntry OUEntry = entry.Children.Add($"OU={OUName}", "OrganizationalUnit"))
                            {
                                OUEntry.Properties["description"].Value = OUDescription;
                                OUEntry.CommitChanges();
                                bl = true;
                            }
                        }
                        else
                        {
                            objLoggerClass.Write("createOU", "OUExists", 0, $"OU '{OUName}' already exists", "OU creation failed due to duplicate"); objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("createOU", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("createOU", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("createOU", "createOU", 0, "create OU in DC", "Exception : " + ex);objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }
        public bool dltOU(string OUName)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        DirectoryEntry ouEntry = entry.Children.Find($"OU={OUName}", "OrganizationalUnit");
                        if (ouEntry != null)
                        {
                            ouEntry.DeleteTree();
                            ouEntry.CommitChanges();
                            bl= true;
                        }
                        else
                        {
                            objLoggerClass.Write("deleteOU", "deleteOU", 0, "delete OU in DC", "OU not found");objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("deleteOU", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("deleteOU", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("deleteOU", "deleteOU", 0, "delete OU in DC", "Exception : " + ex);objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }


        //public bool moveOU(string ouName, string targetOUDN)
        //{
        //    try
        //    {
        //        using (DirectoryEntry rootDSE = new DirectoryEntry($"LDAP://{DomainName}/rootDSE", DomainUsername, DomainUserPassword))
        //        {
        //            string defaultNamingContext = rootDSE.Properties["defaultNamingContext"].Value.ToString();

        //            using (DirectoryEntry entry = new DirectoryEntry($"LDAP://{DomainName}/{defaultNamingContext}", DomainUsername, DomainUserPassword))
        //            {
        //                DirectorySearcher search = new DirectorySearcher(entry);
        //                search.Filter = $"(&(objectClass=organizationalUnit)(name={ouName}))";

        //                SearchResult result = search.FindOne();
        //                if (result != null)
        //                {
        //                    using (DirectoryEntry userEntry = result.GetDirectoryEntry())
        //                    {
        //                        // Connect to the target OU where the user will be moved
        //                        using (DirectoryEntry targetOU = new DirectoryEntry($"LDAP://{DomainName}/{targetOUDN}", DomainUsername, DomainUserPassword))
        //                        {
        //                            userEntry.MoveTo(targetOU);
        //                            userEntry.CommitChanges();
        //                            return true;
        //                        }
        //                    }
        //                }
        //                else
        //                {
        //                    objLoggerClass.Write("MoveUser", "moveUserToOU", 0, "move user to OU in DC", "user not found");
        //                    objLogWriter.WriteLogsFromQueue();
        //                    return false;
        //                }
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        objLoggerClass.Write("MoveUser", "moveUserToOU", 0, " move user to OU in DC", "Exception : " + ex);
        //        objLogWriter.WriteLogsFromQueue();
        //        return false;
        //    }
        //}

        public bool modifyOU(string ouName, string propertyName, string propertyValue)
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
                            srch.Filter = $"(&(objectClass=organizationalUnit)(name={ouName}))";

                                SearchResult result = srch.FindOne();
                                if (result != null)
                                {
                                    using (DirectoryEntry ouEntry = result.GetDirectoryEntry())
                                    {
                                        if (ouEntry.Properties.Contains(propertyName))
                                        {
                                            ouEntry.Properties[propertyName][0] = propertyValue; // Modify existing value
                                        }
                                        else
                                        {
                                            ouEntry.Properties[propertyName].Add(propertyValue); // Add new value if property doesn't exist
                                        }
                                        ouEntry.CommitChanges();
                                        bl = true;
                                    }
                                }
                                else
                                {
                                    objLoggerClass.Write("ModifyOU", "modifyOU", 0, "Modify OU in AD", "OU not found");objLogWriter.WriteLogsFromQueue();
                                }
                            
                        }
                        else
                        {
                            objLoggerClass.Write("ModifyOU", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("ModifyOU", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                    }
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("ModifyOU", "modifyOU", 0, "Modify OU in AD", "Exception : " + ex);objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }


    }
}
