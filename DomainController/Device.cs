using Logger;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomainController
{
    public class Device
    {
        public static string logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\AssertYIT\\DomainController\\";
        LoggerClass objLoggerClass = new LoggerClass("");
        LogWriter objLogWriter = new LogWriter(logFilePath);
        helperMethod objHelper = new helperMethod();
        DirectoryEntry entry;
        DirectorySearcher srch;
        bool isConnect = false;
        string DomainName = "";
        string DomainUsername = "";
        string DomainUserPassword = "";

        public Device(string DomainName, string Username, string Password)
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
                objLoggerClass.Write("Device", "Connect", 0, " can not connect to server ", "exception:" + ex);
                objLogWriter.WriteLogsFromQueue();
            }
            return isConnect;
        }

    //    public bool IsMachineJoinedToDomain()
    //    {
    //        using (var context = new PrincipalContext(ContextType.Machine))
    //        {
    //            return context.ConnectedServer != null;
    //        }
    //    }
        

    //public static List<string> GetADMachines(string domainPath)
    //{
    //    List<string> machines = new List<string>();
    //    DirectoryEntry entry = new DirectoryEntry($"LDAP://{domainPath}");
    //    DirectorySearcher searcher = new DirectorySearcher(entry);

    //    searcher.Filter = "(objectClass=computer)";
    //    searcher.PropertiesToLoad.Add("name");

    //    foreach (SearchResult result in searcher.FindAll())
    //    {
    //        machines.Add(result.Properties["name"][0].ToString());
    //    }
    //    return machines;
    //}

        public bool CreateDevice(string deviceName, string description)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        if (objHelper.isDeviceExists(entry, deviceName))
                        {
                            using (DirectoryEntry newDevice = entry.Children.Add($"CN={deviceName}", "computer"))
                            {
                                newDevice.Properties["sAMAccountName"].Value = deviceName + "$";
                                newDevice.Properties["description"].Value = description;
                                // Commit initial changes
                                newDevice.CommitChanges();

                                newDevice.Properties["userAccountControl"].Value = 0x1000; // Workstation/Server account(enable account)
                                newDevice.CommitChanges();
                                bl = true;
                            }
                        }
                        else
                        {
                            objLoggerClass.Write("CreateDevice", "DeviceExists", 0, $"Device '{deviceName}' already exists", "Device creation failed due to duplicate"); objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("CreateDevice", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("CreateDevice", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("CreateDevice", "CreateDevice", 0, "create device in DC", "exception:" + ex);objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }
        public bool updateDevice(string deviceName, string description)
        { 
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        srch.Filter = string.Format("(&(objectClass=computer)(sAMAccountName={0}$))", deviceName); // Device names typically end with '$' in AD

                        SearchResult result = srch.FindOne();
                        if (result != null)
                        {
                            using (DirectoryEntry deviceEntry = result.GetDirectoryEntry())
                            {
                                deviceEntry.Properties["description"].Value = description;  // Updating device's description
                                deviceEntry.CommitChanges();
                                bl = true;
                            }
                        }
                        else
                        {
                            objLoggerClass.Write("updateDevice", "updateDevice", 0, "update device in DC", "device not found");
                            objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("updateDevice", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("updateDevice", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }   
                
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("updateDevice", "updateDevice", 0, "update device in DC", "exception: " + ex);
                objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }


        public bool enableDevice(string deviceName)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        srch.Filter = string.Format("(&(objectClass=computer)(sAMAccountName={0}$))", deviceName);

                        SearchResult result = srch.FindOne();
                        if (result != null)
                        {
                            DirectoryEntry deviceEntry = result.GetDirectoryEntry();
                            int userAccountControl = (int)deviceEntry.Properties["userAccountControl"].Value;
                            userAccountControl &= ~0x2;   //  userAccountControl |= 0x2;

                            deviceEntry.Properties["userAccountControl"].Value = userAccountControl;
                            deviceEntry.CommitChanges();
                            bl = true;
                        }
                        else
                        {
                            objLoggerClass.Write("enableUser", "enableUser", 0, " enable user in DC", "user not found");
                            objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("enableUser", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("enableUser", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }   
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("enableUser", "enableUser", 0, " enable user in DC", "exception:" + ex);objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }

        public bool disableDevice(string deviceName)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        srch.Filter = string.Format("(&(objectClass=computer)(sAMAccountName={0}$))", deviceName);

                        SearchResult result = srch.FindOne();
                        if (result != null)
                        {
                            DirectoryEntry deviceEntry = result.GetDirectoryEntry();
                            int userAccountControl = (int)deviceEntry.Properties["userAccountControl"].Value;
                            userAccountControl |= 0x2;

                            deviceEntry.Properties["userAccountControl"].Value = userAccountControl;
                            deviceEntry.CommitChanges();
                            bl = true;
                        }
                        else
                        {
                            objLoggerClass.Write("disableUser", "disableUser", 0, " disable user in DC", "user not found");
                            objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("disableUser", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("disableUser", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("disableUser", "disableUser", 0, " disable user in DC", "exception:" + ex); objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }
        public bool deleteDevice(string deviceName)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        srch.Filter = string.Format("(&(objectClass=computer)(sAMAccountName={0}$))", deviceName);

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
                            objLoggerClass.Write("deleteUser", "deleteUser", 0, " delete user in DC", "user not found");objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("deleteUser", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("deleteUser", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("deleteUser", "deleteUser", 0, " delete user in DC", "Exception : " + ex);
                objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }

        public bool moveDeviceToOU(string deviceName, string targetOUDN)
        {
            bool bl = false;
            try
            {
                using (DirectoryEntry rootDSE = new DirectoryEntry($"LDAP://{DomainName}/rootDSE", DomainUsername, DomainUserPassword))
                {
                    string defaultNamingContext = rootDSE.Properties["defaultNamingContext"].Value.ToString();
                    if(Connect())
                    //using (DirectoryEntry entry = new DirectoryEntry($"LDAP://{DomainName}", DomainUsername, DomainUserPassword))
                    {
                        srch.Filter = string.Format("(&(objectClass=computer)(sAMAccountName={0}$))", deviceName);

                        SearchResult result = srch.FindOne();
                        if (result != null)
                        {
                            using (DirectoryEntry deviceEntry = result.GetDirectoryEntry())
                            {
                                using (DirectoryEntry targetOU = new DirectoryEntry($"LDAP://{DomainName}/{targetOUDN},{defaultNamingContext}", DomainUsername, DomainUserPassword))
                                {
                                    if (targetOU != null && targetOU.NativeObject != null)
                                    {
                                        deviceEntry.MoveTo(targetOU);
                                        deviceEntry.CommitChanges();
                                        bl = true;
                                    }
                                    else
                                    {
                                        objLoggerClass.Write("MoveDevice", "MoveDeviceToOU", 0, "move device to OU in DC", "TargetOU not found");
                                        objLogWriter.WriteLogsFromQueue();
                                    }
                                }
                            }
                        }
                        else
                        {
                            objLoggerClass.Write("MoveDevice", "MoveDeviceToOU", 0, "move device to OU in DC", "Device not found");
                            objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("MoveDevice", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                    }
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("MoveDevice", "MoveDeviceToOU", 0, "move device to OU in DC", "Exception: " + ex);
                objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }

    }
}
