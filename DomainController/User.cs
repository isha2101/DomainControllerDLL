using Logger;
using System;
using System.Collections.Generic;
using System.Data;
using System.DirectoryServices;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace DomainController
{
    public class User
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

        public User(string DomainName, string Username, string Password)
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
                objLoggerClass.Write("User", "Connect", 0, " can not connect to server ", "exception:" + ex);
                objLogWriter.WriteLogsFromQueue();
            }
            return isConnect;
        }

        public bool CreateUser(string AccountName, string PrincipleName, string displayName, string description, string password)
        {
            bool bl = false;
            try
            {
                if (Connect()) 
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        if (objHelper.isUserExists(entry, AccountName))
                        {
                            using (DirectoryEntry newUser = entry.Children.Add($"CN={AccountName}", "user"))
                            {
                                newUser.Properties["sAMAccountName"].Value = AccountName;
                                newUser.Properties["userPrincipalName"].Value = PrincipleName;
                                newUser.Properties["displayName"].Value = displayName;
                                newUser.Properties["description"].Value = description;
                                newUser.CommitChanges();

                                if (objHelper.SetPassword(newUser, password))
                                {
                                    newUser.Properties["userAccountControl"].Value = 0x200; // Normal account(Enable the account)
                                    newUser.CommitChanges();
                                }
                                else
                                {
                                    objLoggerClass.Write("CreateUser", "PasswordError", 0, "Failed to set password", "Password operation failed"); objLogWriter.WriteLogsFromQueue();
                                }
                            }
                            bl = true;
                        }
                        else
                        {
                            objLoggerClass.Write("CreateUser", "UserExists", 0, $"User '{AccountName}' already exists", "User creation failed due to duplicate"); objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("CreateUser", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("CreateUser", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("CreatUser", "CreateUser", 0, " create user in DC", "exception:" + ex);
                objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }

        public bool updateUser(string userName, string displayName, string description)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        srch.Filter = string.Format("(&(objectClass=user)(sAMAccountName={0}))", userName);
                        SearchResult result = srch.FindOne();
                        if (result != null)
                        {
                            using (DirectoryEntry userEntry = result.GetDirectoryEntry())
                            {
                                userEntry.Properties["displayName"].Value = displayName;
                                userEntry.Properties["description"].Value = description;
                                userEntry.CommitChanges();
                                bl = true;
                            }
                        }
                        else
                        {
                            objLoggerClass.Write("updateUser", "notFound", 0, " update user in DC", "user not found"); objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("updateUser", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("updateUser", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("updateUser", "updateUser", 0, " update user in DC", "exception:" + ex);objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }

        public bool resetPassword(string userName, string newPassword)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        //DirectorySearcher search = new DirectorySearcher(entry);
                        srch.Filter = string.Format("(&(objectClass=user)(sAMAccountName={0}))", userName);

                        SearchResult result = srch.FindOne();
                        if (result != null)
                        {
                            using (DirectoryEntry userEntry = result.GetDirectoryEntry())
                            {
                                userEntry.Invoke("SetPassword", new object[] { newPassword });
                                userEntry.CommitChanges();
                                objHelper.ValidateUserPassword(userEntry.Path, userName, newPassword);
                                bl = true;
                            }
                        }
                        else
                        {
                            objLoggerClass.Write("resetPWD", "notFound", 0, " reset password of user in DC", "user not found"); objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("resetPWD", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("resetPWD", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("resetPWD", "resetPassword", 0, " reset password of user in DC", "exception:" + ex);objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }

        public bool enableUser(string userName)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        DirectorySearcher search = new DirectorySearcher(entry);
                        search.Filter = $"(sAMAccountName={userName})";

                        SearchResult result = search.FindOne();
                        if (result != null)
                        {
                            DirectoryEntry userEntry = result.GetDirectoryEntry();
                            int userAccountControl = (int)userEntry.Properties["userAccountControl"].Value;
                            userAccountControl &= ~0x2;

                            userEntry.Properties["userAccountControl"].Value = userAccountControl;
                            userEntry.CommitChanges();
                            return true;
                        }
                        else
                        {
                            objLoggerClass.Write("enableUser", "notFound", 0, " enable user in DC", "user not found");objLogWriter.WriteLogsFromQueue();
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

        public bool disableUser(string userName)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        DirectorySearcher search = new DirectorySearcher(entry);
                        search.Filter = $"(sAMAccountName={userName})";

                        SearchResult result = search.FindOne();
                        if (result != null)
                        {
                            DirectoryEntry userEntry = result.GetDirectoryEntry();
                            int userAccountControl = (int)userEntry.Properties["userAccountControl"].Value;
                            userAccountControl |= 0x2;

                            userEntry.Properties["userAccountControl"].Value = userAccountControl;
                            userEntry.CommitChanges();
                            return true;
                        }
                        else
                        {
                            objLoggerClass.Write("disableUser", "notFound", 0, " disable user in DC", "user not found"); objLogWriter.WriteLogsFromQueue();

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
                objLoggerClass.Write("disableUser", "disableUser", 0, " disable user in DC", "exception:" + ex);objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }

        public bool deleteUser(string userName)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        DirectorySearcher search = new DirectorySearcher(entry);
                        search.Filter = $"(sAMAccountName={userName})";

                        SearchResult result = search.FindOne();
                        if (result != null)
                        {
                            DirectoryEntry userEntry = result.GetDirectoryEntry();
                            userEntry.DeleteTree();
                            userEntry.CommitChanges();
                            return true;
                        }
                        else
                        {
                            objLoggerClass.Write("deleteUser", "notFound", 0, "delete user in DC", "user not found"); objLogWriter.WriteLogsFromQueue();
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
                objLoggerClass.Write("deleteUser", "deleteUser", 0, " delete user in DC", "Exception : " + ex);objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }

        public bool moveUserToOU(string userName, string targetOUDN)
        {
            bool bl = false;
            try
            {
                using (DirectoryEntry rootDSE = new DirectoryEntry($"LDAP://{DomainName}/rootDSE", DomainUsername, DomainUserPassword))
                {
                    if(rootDSE != null && rootDSE.NativeObject != null)
                    {
                        string defaultNamingContext = rootDSE.Properties["defaultNamingContext"].Value.ToString();
                        if (Connect())
                        {
                            if (objHelper.AreCredentialValid(entry))
                            {
                                DirectorySearcher search = new DirectorySearcher(entry);
                                search.Filter = $"(sAMAccountName={userName})";

                                SearchResult result = search.FindOne();
                                if (result != null)
                                {
                                    using (DirectoryEntry userEntry = result.GetDirectoryEntry())
                                    {
                                        using (DirectoryEntry targetOU = new DirectoryEntry($"LDAP://{DomainName}/{targetOUDN},{defaultNamingContext}", DomainUsername, DomainUserPassword))
                                        {
                                            if (targetOU != null && targetOU.NativeObject != null)
                                            {
                                                userEntry.MoveTo(targetOU);
                                                userEntry.CommitChanges();
                                                bl= true;
                                            }
                                            else
                                            {
                                                objLoggerClass.Write("MoveUser", "notFound", 0, "move user to OU in DC", "target OU not found"); objLogWriter.WriteLogsFromQueue();

                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    objLoggerClass.Write("MoveUser", "notFound", 0, "move user to OU in DC", "user not found"); objLogWriter.WriteLogsFromQueue();

                                }
                            }
                            else
                            {
                                objLoggerClass.Write("MoveUser", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();

                            }
                        }
                        else
                        {
                            objLoggerClass.Write("MoveUser", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("MoveUser", "moveUserToOU", 0, "Failed to connect to rootDSA", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                    }
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("MoveUser", "moveUserToOU", 0, "move user to OU in DC", "Exception : " + ex);objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }
        
        public string getUserPwdStatus(string userName)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        using (DirectorySearcher search = new DirectorySearcher(entry))
                        {
                            search.Filter = $"(sAMAccountName={userName})";
                            SearchResult result = search.FindOne();

                            if (result != null)
                            {
                                DirectoryEntry userEntry = result.GetDirectoryEntry();

                                long pwdLastSet = 0;
                                if (userEntry.Properties["pwdLastSet"].Value != null)
                                {
                                    pwdLastSet = objHelper.ConvertLargeIntegerToLong(userEntry.Properties["pwdLastSet"].Value);
                                }
                                DateTime passwordLastSetDate = DateTime.FromFileTimeUtc(pwdLastSet);

                                long expiresLong = 0;
                                if (userEntry.Properties["accountExpires"].Value != null)
                                {
                                    expiresLong = objHelper.ConvertLargeIntegerToLong(userEntry.Properties["accountExpires"].Value);
                                }
                                string accountExpires = null;
                                if (expiresLong == 0 || expiresLong == long.MaxValue)
                                {
                                    accountExpires = "Account does not expire";
                                }
                                else
                                {
                                    try
                                    {
                                        DateTime accountExpiresDate = DateTime.FromFileTimeUtc(expiresLong);
                                        accountExpires = accountExpiresDate.ToString("yyyy-MM-dd HH:mm:ss");
                                    }
                                    catch (ArgumentOutOfRangeException)
                                    {
                                        accountExpires = null; // Handle invalid file time values gracefully
                                    }
                                }
                                return $"Password Last Set: {passwordLastSetDate}\n" +
                                        $"Account Expiration Date: {(expiresLong == 0 ? "Never" : accountExpires.ToString())}\n";
                            }
                            else return "User not found";
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("pwdStatus", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                        return "Validation Failed";
                    }
                }
                else
                {
                    objLoggerClass.Write("pwdStatus", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                    return "Connection not established";
                }
                
            }
            catch (Exception ex) 
            { return "Error: " + ex.Message; }
        }

        public DataTable getPwdPolicy()
        {
            DataTable passwordPolicyTable = new DataTable();
            passwordPolicyTable.Columns.Add("Policy Name", typeof(string));
            passwordPolicyTable.Columns.Add("Value", typeof(string));
            try
            {
                using (DirectoryEntry entry = new DirectoryEntry($"LDAP://{DomainName}/rootDSE", DomainUsername, DomainUserPassword))
                {
                    string defaultNamingContext = entry.Properties["defaultNamingContext"].Value.ToString();
                    using (DirectoryEntry domainEntry = new DirectoryEntry($"LDAP://{DomainName}/{defaultNamingContext}", DomainUsername, DomainUserPassword))
                    {
                        int minPasswordLength = (int)domainEntry.Properties["minPwdLength"].Value;
                        int maxPasswordAge = objHelper.ConvertLargeIntegerToDays(domainEntry.Properties["maxPwdAge"].Value);
                        int minPasswordAge = objHelper.ConvertLargeIntegerToDays(domainEntry.Properties["minPwdAge"].Value);
                        bool passwordComplexity = Convert.ToBoolean(domainEntry.Properties["pwdProperties"].Value);

                        passwordPolicyTable.Rows.Add("Minimum Password Length", minPasswordLength.ToString());
                        passwordPolicyTable.Rows.Add("Maximum Password Age", $"{maxPasswordAge} days");
                        passwordPolicyTable.Rows.Add("Minimum Password Age", $"{minPasswordAge} days");
                        passwordPolicyTable.Rows.Add("Password Complexity Required", passwordComplexity ? "Yes" : "No");

                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error retrieving password policy: " + ex.Message);
            }
            return passwordPolicyTable;
        }

        public void setCustomPwd(string userName)
        {
            try
            {
                using (DirectoryEntry entry = new DirectoryEntry($"LDAP://{DomainName}", DomainUsername, DomainUserPassword))
                {
                    DirectorySearcher search = new DirectorySearcher(entry);
                    search.Filter = $"(sAMAccountName={userName})";

                    SearchResult result = search.FindOne();
                    if (result != null)
                    {
                        using (DirectoryEntry userEntry = result.GetDirectoryEntry())
                        {
                            // set user account control flag
                            int userAccountControl = (int)userEntry.Properties["userAccountControl"].Value;

                            userAccountControl |= 0x10000; //0x10000 is the flag for Password Never Expires

                            userEntry.Properties["pwdLastSet"].Value = 0;

                            userEntry.Properties["userAccountControl"].Value = userAccountControl;
                            userEntry.CommitChanges();
                        }
                    }
                    else
                    {
                        Console.WriteLine("User not found.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error retrieving password policy: " + ex.Message);
            }
        }
    

    }
}
