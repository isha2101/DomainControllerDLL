using Logger;
using Microsoft.Win32;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using OwnYITCommon;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.IO;
using System.Linq;
using System.Management.Automation;
using System.Net;
using System.Security.AccessControl;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using static System.Net.WebRequestMethods;

namespace DomainController
{
    public class List
    {
        public static string logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\AssertYIT\\DomainController\\";
        LoggerClass objLoggerClass = new LoggerClass("");
        LogWriter objLogWriter = new LogWriter(logFilePath);
        helperMethod objHelper = new helperMethod();
        DataTableConversion objDT = new DataTableConversion();
        DirectoryEntry entry;
        DirectorySearcher srch;
        bool isConnect = false;
        string DomainName = "";
        string DomainUsername = "";
        string DomainUserPassword = "";

        public List(string DomainName, string Username, string Password)
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

                srch.FindOne();

                isConnect = true;
            }
            catch (Exception ex)
            {
                isConnect = false;
                objLoggerClass.Write("List", "Connect", 0, " can not connect to server ", "exception:" + ex); objLogWriter.WriteLogsFromQueue();
            }
            return isConnect;
        }

        /// <summary>
        /// Retrieves basic information about the current Active Directory domain.
        /// </summary>
        private DomainInfo GetDomainInfo()
        {
            var info = new DomainInfo();

            try
            {
                // Connect to the rootDSE object to fetch default naming context
                using (DirectoryEntry rootDSE = new DirectoryEntry($"LDAP://{DomainName}/rootDSE", DomainUsername, DomainUserPassword))
                {
                    string defaultNamingContext = rootDSE.Properties["defaultNamingContext"].Value.ToString();
                    info.DomainHost = defaultNamingContext;

                    using (DirectoryEntry domainEntry = new DirectoryEntry($"LDAP://{DomainName}/{defaultNamingContext}", DomainUsername, DomainUserPassword))
                    {
                        info.DomainGUID = domainEntry.Guid.ToString();

                        if (domainEntry.Properties.Contains("objectSid"))
                        {
                            byte[] sidBytes = (byte[])domainEntry.Properties["objectSid"].Value;
                            var sid = new System.Security.Principal.SecurityIdentifier(sidBytes, 0);
                            info.DomainSID = sid.Value;
                        }
                        else
                        {
                            info.DomainSID = "N/A";
                        }
                    }
                    // Search for a Domain Controller object using userAccountControl filter for DCs
                    using (DirectoryEntry searchRoot = new DirectoryEntry($"LDAP://{DomainName}/{defaultNamingContext}", DomainUsername, DomainUserPassword))
                    using (DirectorySearcher dcSearcher = new DirectorySearcher(searchRoot))
                    {
                        dcSearcher.Filter = "(userAccountControl:1.2.840.113556.1.4.803:=532480)";
                        dcSearcher.PropertiesToLoad.Add("dNSHostName");

                        SearchResult result = dcSearcher.FindOne();
                        if (result != null && result.Properties.Contains("dNSHostName"))
                        {
                            info.DomainControllerName = result.Properties["dNSHostName"][0].ToString();
                            info.IPAddress = DomainName;
                        }
                        else
                        {
                            info.DomainControllerName = "N/A";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                info.DomainHost = "Error fetching domain";
                info.DomainGUID = "N/A";
                info.DomainSID = "N/A";
                info.DomainControllerName = "N/A";
                info.IPAddress = "N/A";
                Console.WriteLine($"Error: {ex.Message}");
            }

            return info;
        }
        public bool IsInternetBlockedByGPO(string username, string domain)
        {
            using (DirectorySearcher searcher = new DirectorySearcher(new DirectoryEntry($"LDAP://{domain}")))
            {
                searcher.Filter = $"(&(objectClass=user)(samaccountname={username}))";
                SearchResult result = searcher.FindOne();

                if (result != null)
                {
                    DirectoryEntry userEntry = result.GetDirectoryEntry();

                    if (userEntry.Properties.Contains("msDS-ResultantPSO"))
                    {
                        string policy = userEntry.Properties["msDS-ResultantPSO"][0].ToString();
                        if (policy.Contains("NoInternetPolicy"))
                        {
                            return true; // User has GPO blocking internet
                        }
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Retrieves a list of users who are members of high-privilege Active Directory groups
        /// (e.g., Domain Admins, Enterprise Admins, etc.). For each user, it checks:
        /// - Their SID and GUID
        /// - Whether they have potential internet access (based on proxy address or network test)
        /// - Whether they are part of any internet-restricted groups
        /// The resulting data is returned as a DataTable with group-user linkage and internet access status.
        /// </summary>
        public DataTable GetHighPrivilegeUser()
        {
            // List of high-privilege groups
            List<string> highPrivilegeGroups = new List<string>
    {
        "Domain Admins",
        "Enterprise Admins",
        "Administrators",
        "Account Operators",
        "Schema Admins"
    };

            // Create a DataTable to store the Group and User linkage
            DataTable dtProperty = new DataTable();
            dtProperty.Columns.Add("Group");
            dtProperty.Columns.Add("User");
            dtProperty.Columns.Add("SID");
            dtProperty.Columns.Add("GUID");
            dtProperty.Columns.Add("InternetAccess");

            try
            {
                // Proceed if the main DirectoryEntry is initialized
                if (entry != null)
                {
                    srch.Filter = "(&(objectClass=group))";
                    SearchResultCollection groupResults = srch.FindAll();

                    if (groupResults != null)
                    {
                        foreach (SearchResult srGroup in groupResults)
                        {
                            string groupName = "";
                            string ObjectSID = "";
                            string ObjectGUID = "";
                            ResultPropertyCollection groupProps = srGroup.Properties;

                            // Retrieve the group name (samaccountname or name)
                            foreach (string groupPropName in groupProps.PropertyNames)
                            {
                                if (groupPropName == "samaccountname")
                                {
                                    groupName = groupProps[groupPropName][0].ToString();
                                    break;
                                }
                            }

                            // Process only high-privilege groups
                            if (highPrivilegeGroups.Contains(groupName, StringComparer.OrdinalIgnoreCase))
                            {
                                // Retrieve the members of the group
                                if (groupProps["member"] != null)
                                {
                                    foreach (var member in groupProps["member"])
                                    {
                                        string userDN = member.ToString();

                                        try
                                        {
                                            using (DirectoryEntry userEntry = new DirectoryEntry($"LDAP://{DomainName}/{userDN}", DomainUsername, DomainUserPassword))
                                            {
                                                DirectorySearcher userSearcher = new DirectorySearcher(userEntry);
                                                userSearcher.Filter = "(&(objectCategory=User)(objectClass=person))";
                                                SearchResult userResult = userSearcher.FindOne();

                                                if (userResult != null)
                                                {
                                                    string userName = "";
                                                    bool hasInternetAccess = false;
                                                    ResultPropertyCollection userProps = userResult.Properties;

                                                    foreach (string userPropName in userProps.PropertyNames)
                                                    {
                                                        if (userPropName == "samaccountname")
                                                        {
                                                            userName = userProps[userPropName][0].ToString();
                                                        }

                                                        // Check if user has external email in proxyAddresses
                                                        if (userPropName == "proxyAddresses" && userProps[userPropName].Count > 0)
                                                        {
                                                            foreach (var address in userProps[userPropName])
                                                            {
                                                                string email = address.ToString();
                                                                if (email.StartsWith("SMTP:", StringComparison.OrdinalIgnoreCase))
                                                                {
                                                                    string emailDomain = email.Split(':')[1].Split('@').Last();
                                                                    if (!emailDomain.EndsWith("company.com", StringComparison.OrdinalIgnoreCase)) // Change company domain
                                                                    {
                                                                        hasInternetAccess = true;
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }

                                                    // Check if user is in an Internet-Blocked AD Group
                                                    List<string> internetRestrictedGroups = new List<string> { "NoInternetUsers", "RestrictedAccess" };
                                                    foreach (var group in userProps["memberOf"])
                                                    {
                                                        string userGroup = group.ToString().Split(',')[0].Replace("CN=", "");
                                                        if (internetRestrictedGroups.Contains(userGroup, StringComparer.OrdinalIgnoreCase))
                                                        {
                                                            hasInternetAccess = false;
                                                            break;
                                                        }
                                                    }


                                                    // Check Internet Connectivity using PowerShell
                                                    bool canAccessInternet = false;
                                                    try
                                                    {
                                                        using (Process process = new Process())
                                                        {
                                                            process.StartInfo.FileName = "powershell.exe";
                                                            process.StartInfo.Arguments = "-Command \"Test-NetConnection google.com -Port 443 | Select-Object -ExpandProperty TcpTestSucceeded\"";
                                                            process.StartInfo.RedirectStandardOutput = true;
                                                            process.StartInfo.UseShellExecute = false;
                                                            process.StartInfo.CreateNoWindow = true;

                                                            process.Start();
                                                            string output = process.StandardOutput.ReadToEnd();
                                                            process.WaitForExit();

                                                            canAccessInternet = output.Trim().Equals("True", StringComparison.OrdinalIgnoreCase);
                                                        }
                                                    }
                                                    catch (Exception)
                                                    {
                                                        canAccessInternet = false;
                                                    }

                                                    hasInternetAccess = hasInternetAccess || canAccessInternet;

                                                    byte[] sidBytes = (byte[])userResult.Properties["objectSid"][0];
                                                    System.Security.Principal.SecurityIdentifier sid = new System.Security.Principal.SecurityIdentifier(sidBytes, 0);
                                                    ObjectSID = sid.Value;
                                                    Guid groupGUID = new Guid((byte[])srGroup.Properties["objectGUID"][0]);
                                                    ObjectGUID = groupGUID.ToString();

                                                    // Add the Group-User linkage to the DataTable
                                                    DataRow drLinkage = dtProperty.NewRow();
                                                    drLinkage["Group"] = groupName;
                                                    drLinkage["User"] = userName;
                                                    drLinkage["SID"] = ObjectSID;
                                                    drLinkage["GUID"] = ObjectGUID;
                                                    drLinkage["InternetAccess"] = hasInternetAccess ? "Yes" : "No";
                                                    dtProperty.Rows.Add(drLinkage);
                                                }
                                            }
                                        }
                                        catch (Exception userEx)
                                        {
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("linkage", "HighPrivilegeGroupUserLinkage", 0, " get linkage of high-privilege Group and User", "exception:" + ex);
                objLogWriter.WriteLogsFromQueue();
            }
            return dtProperty;
        }

        /// <summary>
        /// Retrieves a list of domain user accounts that have the "Password Never Expires" flag set
        /// and are not disabled. For each such user, gathers detailed information in a structured DataTable.
        /// </summary>
        public DataTable GetUserPwd()
        {
            string UserName = "";
            string LongName = "";
            string ObjectGUID = "";
            string ObjectSID = "";
            string CreatedDate = "";
            string ModifiedDate = "";
            string DomainHost = "";
            string DomainGUID = "";
            string DomainSID = "";
            string DomainControllerName = "";
            string IPAddress = "";
            //string passwordLastSetDate = "";

            // Create a DataTable to store user information
            DataTable dtUsers = new DataTable();
            dtUsers.Columns.Add("UserName");
            dtUsers.Columns.Add("LongName");
            dtUsers.Columns.Add("ObjectGUID");
            dtUsers.Columns.Add("ObjectSID");
            dtUsers.Columns.Add("CreatedDate");
            dtUsers.Columns.Add("ModifiedDate");
            dtUsers.Columns.Add("passwordLastSetDate");
            dtUsers.Columns.Add("Domainname");
            dtUsers.Columns.Add("DomainGUID");
            dtUsers.Columns.Add("DomainSID");
            dtUsers.Columns.Add("DomainControllerName");
            dtUsers.Columns.Add("IPAddress");

            // Get domain-related information using helper method
            var domainInfo = GetDomainInfo();
            DomainHost = domainInfo.DomainHost;
            DomainGUID = domainInfo.DomainGUID;
            DomainSID = domainInfo.DomainSID;
            DomainControllerName = domainInfo.DomainControllerName;
            IPAddress = domainInfo.IPAddress;

            // Ensure DirectoryEntry is initialized
            if (entry != null)
            {
                // LDAP filter to retrieve user accounts marked as "Password Never Expires"
                srch.Filter = "(&(objectClass=user)(objectCategory=person)(userAccountControl:1.2.840.113556.1.4.803:=65536))";
                Thread.Sleep(1000);
                try
                {
                    SearchResultCollection srcColl = srch.FindAll();
                    if (srcColl != null)
                    {
                        foreach (SearchResult srUser in srcColl)
                        {
                            if (srUser.Properties.Contains("cn"))
                            {
                                UserName = srUser.Properties["cn"][0].ToString();
                            }
                            long pwdLastSet = 0;
                            if (srUser.Properties["pwdLastSet"][0] != null)
                            {
                                pwdLastSet = (long)srUser.Properties["pwdLastSet"][0];
                               
                            }
                            DateTime passwordLastSetDate = DateTime.FromFileTimeUtc(pwdLastSet);
                                
                            Guid userGUID = new Guid((byte[])srUser.Properties["objectGUID"][0]);
                            byte[] sidBytes = (byte[])srUser.Properties["objectSid"][0];
                            System.Security.Principal.SecurityIdentifier sid = new System.Security.Principal.SecurityIdentifier(sidBytes, 0);
                            ObjectSID = sid.Value;

                            if (srUser.Properties.Contains("whenCreated"))
                            {
                                CreatedDate = srUser.Properties["whenCreated"][0].ToString();
                            }
                            if (srUser.Properties.Contains("whenChanged"))
                            {
                                ModifiedDate = srUser.Properties["whenChanged"][0].ToString();
                            }
                            string dnParts = srUser.Path.Split('/').Last();
                            LongName = objHelper.getOULong(srUser.Path.Split(','));

                            // Check the account status
                            if (srUser.Properties.Contains("userAccountControl"))
                            {
                                int userAccountControl = (int)srUser.Properties["userAccountControl"][0];
                                bool isPasswordNeverExpires = (userAccountControl & 0x10000) != 0;
                                bool isDisabled = (userAccountControl & 0x2) != 0;

                                // Include only active accounts with "Password Never Expires" set
                                if (isPasswordNeverExpires && !isDisabled )
                                {
                                    DataRow dr = dtUsers.NewRow();
                                    dr["UserName"] = UserName;
                                    dr["LongName"] = dnParts;
                                    dr["ObjectGUID"] = userGUID.ToString();
                                    dr["ObjectSID"] = ObjectSID;
                                    dr["CreatedDate"] = CreatedDate;
                                    dr["ModifiedDate"] = ModifiedDate;
                                    dr["passwordLastSetDate"] = passwordLastSetDate;
                                    dr["Domainname"] = DomainHost;
                                    dr["DomainGUID"] = DomainGUID;
                                    dr["DomainSID"] = DomainSID;
                                    dr["DomainControllerName"] = DomainControllerName;
                                    dr["IPAddress"] = IPAddress;
                                    dtUsers.Rows.Add(dr);
                                }
                            }
                        }

                    }
                }
                catch (Exception ex)
                {
                    DataRow errorRow = dtUsers.NewRow();
                    errorRow["UserName"] = "Error";
                    errorRow["LongName"] = "N/A";
                    errorRow["ObjectGUID"] = "N/A";
                    errorRow["ObjectSID"] = "N/A";
                    errorRow["CreatedDate"] = "N/A";
                    errorRow["ModifiedDate"] = "N/A";
                    dtUsers.Rows.Add(errorRow);
                }
            }
            return dtUsers; // Return the final DataTable of qualified users
        }

        /// <summary>
        /// Retrieves a list of service accounts from the domain. 
        /// Service accounts are identified based on specific flags such as `PasswordNeverExpires`, 
        /// and account is must be enable,
        /// and naming patterns (e.g., starting with "svc_") or their location in the "Managed Service Accounts" container.
        /// Returns detailed info like GUID, SID, timestamps, and domain-related data in a DataTable.
        /// </summary>
        public DataTable GetServiceAccountList()
        {
            string name = "";
            string UserName = "";
            string LongName = "";
            string ObjectGUID = "";
            string ObjectSID = "";
            string CreatedDate = "";
            string ModifiedDate = "";
            string DomainHost = "";
            string DomainGUID = "";
            string DomainSID = "";
            string DomainControllerName = "";
            string IPAddress = "";
            // Define schema for DataTable to store results
            DataTable dtServiceAccounts = new DataTable();
            dtServiceAccounts.Columns.Add("name");
            dtServiceAccounts.Columns.Add("UserName");
            dtServiceAccounts.Columns.Add("LongName");
            dtServiceAccounts.Columns.Add("ObjectGUID");
            dtServiceAccounts.Columns.Add("ObjectSID");
            dtServiceAccounts.Columns.Add("CreatedDate");
            dtServiceAccounts.Columns.Add("ModifiedDate");
            dtServiceAccounts.Columns.Add("PasswordLastSetDate");
            dtServiceAccounts.Columns.Add("Domainname");
            dtServiceAccounts.Columns.Add("DomainGUID");
            dtServiceAccounts.Columns.Add("DomainSID");
            dtServiceAccounts.Columns.Add("DomainControllerName");
            dtServiceAccounts.Columns.Add("IPAddress");
            var domainInfo = GetDomainInfo();
            DomainHost = domainInfo.DomainHost;
            DomainGUID = domainInfo.DomainGUID;
            DomainSID = domainInfo.DomainSID;
            DomainControllerName = domainInfo.DomainControllerName;
            IPAddress = domainInfo.IPAddress;

            if (entry != null)
            {
                // Search filter: users with SERVICE_ACCOUNT flag (0x10000)
                srch.Filter = "(&(objectClass=user)(objectCategory=person)(userAccountControl:1.2.840.113556.1.4.803:=65536))";
                Thread.Sleep(1000);

                try
                {
                    SearchResultCollection srcColl = srch.FindAll();
                    if (srcColl != null)
                    {
                        foreach (SearchResult srUser in srcColl)
                        {
                            if (srUser.Properties.Contains("userprincipalname"))
                                UserName = srUser.Properties["userprincipalname"][0].ToString();
                            if (srUser.Properties.Contains("cn"))
                                name = srUser.Properties["cn"][0].ToString();

                            long pwdLastSet = srUser.Properties.Contains("pwdLastSet") ? (long)srUser.Properties["pwdLastSet"][0] : 0;
                            DateTime passwordLastSetDate = DateTime.FromFileTimeUtc(pwdLastSet);

                            Guid userGUID = new Guid((byte[])srUser.Properties["objectGUID"][0]);
                            byte[] sidBytes = (byte[])srUser.Properties["objectSid"][0];
                            System.Security.Principal.SecurityIdentifier sid = new System.Security.Principal.SecurityIdentifier(sidBytes, 0);
                            ObjectSID = sid.Value;

                            CreatedDate = srUser.Properties.Contains("whenCreated") ? srUser.Properties["whenCreated"][0].ToString() : "N/A";
                            ModifiedDate = srUser.Properties.Contains("whenChanged") ? srUser.Properties["whenChanged"][0].ToString() : "N/A";
                            string dnParts = srUser.Path.Split('/').Last();
                            LongName = objHelper.getOULong(srUser.Path.Split(','));

                            int userAccountControl = (int)srUser.Properties["userAccountControl"][0];
                            bool isPasswordNeverExpires = (userAccountControl & 0x10000) != 0;
                            bool isDisabled = (userAccountControl & 0x2) != 0;

                            // Validate account as service account:
                            // Must have PasswordNeverExpires, not disabled, not computer account
                            // OR be in known service account patterns or containers
                            if (isPasswordNeverExpires && !isDisabled && (userAccountControl & 0x200) == 0
                        || (UserName.StartsWith("svc_")
                            || srUser.Path.Contains("CN=Managed Service Accounts")
                            || srUser.Path.Contains("OU=Managed Service Accounts"))
                        && !UserName.EndsWith("$"))
                            {
                                DataRow dr = dtServiceAccounts.NewRow();
                                dr["name"] = name;
                                dr["UserName"] = UserName;
                                dr["LongName"] = dnParts;
                                dr["ObjectGUID"] = userGUID.ToString();
                                dr["ObjectSID"] = ObjectSID;
                                dr["CreatedDate"] = CreatedDate;
                                dr["ModifiedDate"] = ModifiedDate;
                                dr["PasswordLastSetDate"] = passwordLastSetDate;
                                dr["Domainname"] = DomainHost;
                                dr["DomainGUID"] = DomainGUID;
                                dr["DomainSID"] = DomainSID;
                                dr["DomainControllerName"] = DomainControllerName;
                                dr["IPAddress"] = IPAddress;
                                dtServiceAccounts.Rows.Add(dr);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    DataRow errorRow = dtServiceAccounts.NewRow();
                    errorRow["UserName"] = "Error";
                    dtServiceAccounts.Rows.Add(errorRow);
                }
            }
            return dtServiceAccounts;// Return the full table of service accounts
        }

        /// <summary>
        /// Retrieves all Organizational Units (OUs) from Active Directory and returns their details in a DataTable.
        /// Also includes domain-related metadata such as Domain GUID, SID, and Domain Controller information.
        /// </summary>
        public DataTable GetAllOU()
        {
            // Variable declarations for OU and domain metadata
            string OUNAme = "";
            string LongName = "";
            string ObjectGUID = "";
            string CreatedDate = "";
            string ModifiedDate = "";
            string Status = "";
            string DomainHost = "";
            string DomainGUID = "";
            string DomainSID = "";
            string DomainControllerName = "";
            string IPAddress = "";
            // Create and structure the result DataTable
            DataTable dtOU = new DataTable();
            dtOU.Columns.Add("OUName");
            dtOU.Columns.Add("LongName");
            dtOU.Columns.Add("ObjectGUID");
            dtOU.Columns.Add("CreatedDate");
            dtOU.Columns.Add("ModifiedDate");
            dtOU.Columns.Add("Status");
            dtOU.Columns.Add("Domainname");
            dtOU.Columns.Add("DomainGUID");
            dtOU.Columns.Add("DomainSID");
            dtOU.Columns.Add("DomainControllerName");
            dtOU.Columns.Add("IPAddress");
            // Fetch domain information using helper method
            var domainInfo = GetDomainInfo();
            DomainHost = domainInfo.DomainHost;
            DomainGUID = domainInfo.DomainGUID;
            DomainSID = domainInfo.DomainSID;
            DomainControllerName = domainInfo.DomainControllerName;
            IPAddress = domainInfo.IPAddress;

            // If the DirectoryEntry is valid, search for OUs
            if (entry != null)
            {
                srch.Filter = "(&(objectClass=organizationalUnit))";// Filter only OUs
                Thread.Sleep(1000);
                try
                {
                    SearchResultCollection srcColl = srch.FindAll();
                    if (srcColl != null)
                    {
                        foreach (SearchResult srOU in srcColl)
                        {
                            string ouDN = srOU.Path.Split('/').Last();  // Extract OU distinguished name
                            string ouLonName = objHelper.getOULongName(ouDN);
                            Guid ouGUID = new Guid((byte[])srOU.Properties["objectGUID"][0]);
                            CreatedDate = srOU.Properties["whenCreated"][0].ToString();
                            ModifiedDate = srOU.Properties["whenChanged"][0].ToString();
                           
                            if (srOU.Properties.Contains("userAccountControl"))
                            {
                                int userAccountControl = (int)srOU.Properties["userAccountControl"][0];
                                bool isAccountDisabled = (userAccountControl & 0x2) != 0; // Check if the 'Account Disabled' bit is set

                                Status = isAccountDisabled ? "Disabled" : "Enabled";
                            }

                            // Loop through OU properties and add to DataTable
                            ResultPropertyCollection rpc = srOU.Properties;
                            foreach (string rp in rpc.PropertyNames)
                            {
                                if (rp == "name")
                                {
                                    OUNAme = srOU.Properties[rp][0].ToString();                                 
                                    LongName = ouLonName;
                                    ObjectGUID = ouGUID.ToString();  

                                    DataRow dr = dtOU.NewRow();
                                    dr["OUName"] = OUNAme;
                                    dr["LongName"] = LongName;
                                    dr["ObjectGUID"] = ObjectGUID;
                                    dr["CreatedDate"] = CreatedDate;
                                    dr["ModifiedDate"] = ModifiedDate;
                                    dr["Status"] = Status;
                                    dr["Domainname"] = DomainHost;
                                    dr["DomainGUID"] = DomainGUID;
                                    dr["DomainSID"] = DomainSID;
                                    dr["DomainControllerName"] = DomainControllerName;
                                    dr["IPAddress"] = IPAddress;

                                    dtOU.Rows.Add(dr);// Add row to DataTable
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    DataRow errorRow = dtOU.NewRow();
                    errorRow["OUName"] = "Error";
                    errorRow["LongName"] = ex.Message;
                    errorRow["ObjectGUID"] = "N/A";
                    errorRow["CreatedDate"] = "N/A";
                    errorRow["ModifiedDate"] = "N/A";
                    errorRow["Domainame"] = "Error fetching domain";
                    errorRow["DomainGUID"] = "N/A";
                    errorRow["DomainSID"] = "N/A";
                    dtOU.Rows.Add(errorRow);
                }
            }
            return dtOU;// Return the populated DataTable
        }

        /// <summary>
        /// Retrieves all Users from Active Directory and returns their details in a DataTable.
        /// Also includes domain-related metadata such as Domain GUID, SID, and Domain Controller information.
        /// </summary>
        public DataTable GetAllUsers()
        {
            // Variable declarations for User and domain metadata
            string internetAccess = "";
            string UserName = "";
            string LongName = "";
            string ObjectGUID = "";
            string ObjectSID = "";
            string CreatedDate = "";
            string ModifiedDate = "";
            string Status = "";
            string DomainHost = "";
            string DomainGUID = "";
            string DomainSID = "";
            string DomainControllerName = "";
            string IPAddress = "";

            // Create a DataTable to store user information
            DataTable dtUsers = new DataTable();
            dtUsers.Columns.Add("internetAccess");
            dtUsers.Columns.Add("UserName");
            dtUsers.Columns.Add("LongName");
            dtUsers.Columns.Add("ObjectGUID");
            dtUsers.Columns.Add("ObjectSID");
            dtUsers.Columns.Add("CreatedDate");
            dtUsers.Columns.Add("ModifiedDate");
            dtUsers.Columns.Add("Status");
            dtUsers.Columns.Add("Domainname");
            dtUsers.Columns.Add("DomainGUID");
            dtUsers.Columns.Add("DomainSID");
            dtUsers.Columns.Add("DomainControllerName");
            dtUsers.Columns.Add("IPAddress");

            // Fetch domain information using helper method
            var domainInfo = GetDomainInfo();
            DomainHost = domainInfo.DomainHost;
            DomainGUID = domainInfo.DomainGUID;
            DomainSID = domainInfo.DomainSID;
            DomainControllerName = domainInfo.DomainControllerName;
            IPAddress = domainInfo.IPAddress;

            // If the DirectoryEntry is valid, search for Users
            if (entry != null)
            {
                srch.Filter = "(&(objectClass=user)(objectCategory=person))"; // Filter only Users
                Thread.Sleep(1000);
                try
                {
                    SearchResultCollection srcColl = srch.FindAll();
                    if (srcColl != null)
                    {
                        // Loop through Users properties and add to DataTable
                        foreach (SearchResult srUser in srcColl)
                        {
                            if (srUser.Properties.Contains("msDS-ResultantPSO"))
                            {
                                string policy = srUser.Properties["msDS-ResultantPSO"][0].ToString();
                                if (policy.Contains("internet"))
                                {
                                    internetAccess = "yes";
                                }
                                else
                                {
                                    internetAccess = "no";
                                }
                            }
                            if (srUser.Properties.Contains("cn"))
                            {
                                UserName = srUser.Properties["cn"][0].ToString();
                            }
                            Guid userGUID = new Guid((byte[])srUser.Properties["objectGUID"][0]);
                            byte[] sidBytes = (byte[])srUser.Properties["objectSid"][0];
                            System.Security.Principal.SecurityIdentifier sid = new System.Security.Principal.SecurityIdentifier(sidBytes, 0);
                            ObjectSID = sid.Value;

                            if (srUser.Properties.Contains("whenCreated"))
                            {
                                CreatedDate = srUser.Properties["whenCreated"][0].ToString();
                            }
                            if (srUser.Properties.Contains("whenChanged"))
                            {
                                ModifiedDate = srUser.Properties["whenChanged"][0].ToString();
                            }
                            string dnParts = srUser.Path.Split('/').Last();
                            LongName = objHelper.getOULongName(dnParts);

                            // Check the account status
                            if (srUser.Properties.Contains("userAccountControl"))
                            {
                                int userAccountControl = (int)srUser.Properties["userAccountControl"][0];
                                bool isAccountDisabled = (userAccountControl & 0x2) != 0; // Check if the 'Account Disabled' bit is set

                                Status = isAccountDisabled ? "Disabled" : "Enabled";
                            }
                         

                            DataRow dr = dtUsers.NewRow();
                            dr["internetAccess"] = internetAccess;
                            dr["UserName"] = UserName;
                            dr["LongName"] = LongName;
                            dr["ObjectGUID"] = userGUID.ToString();  // Convert GUID to string
                            dr["ObjectSID"] = ObjectSID;  // Store SID as string
                            dr["CreatedDate"] = CreatedDate;
                            dr["ModifiedDate"] = ModifiedDate;
                            dr["Status"] = Status;
                            dr["Domainname"] = DomainHost;
                            dr["DomainGUID"] = DomainGUID;
                            dr["DomainSID"] = DomainSID;
                            dr["DomainControllerName"] = DomainControllerName;
                            dr["IPAddress"] = IPAddress;
                            dtUsers.Rows.Add(dr);
                        }
                    }
                }
                catch (Exception ex)
                {
                    DataRow errorRow = dtUsers.NewRow();
                    errorRow["UserName"] = "Error";
                    errorRow["LongName"] = "N/A";
                    errorRow["ObjectGUID"] = "N/A";
                    errorRow["ObjectSID"] = "N/A";
                    errorRow["CreatedDate"] = "N/A";
                    errorRow["ModifiedDate"] = "N/A";
                    errorRow["Status"] = "N/A";
                    errorRow["Domainname"] = "Error fetching domain";
                    errorRow["DomainGUID"] = "N/A";
                    errorRow["DomainSID"] = "N/A";
                    dtUsers.Rows.Add(errorRow);// Add row to DataTable
                }
            }
            return dtUsers;// Return the populated DataTable
        }

        /// <summary>
        /// Retrieves all Devices from Active Directory and returns their details in a DataTable.
        /// Also includes domain-related metadata such as Domain GUID, SID, and Domain Controller information.
        /// </summary>
        public DataTable GetAllDevices()
        {// Variable declarations for Devices and domain metadata
            string DeviceName = "";
            string ObjectGUID = "";
            string ObjectSID = "";
            string CreatedDate = "";
            string ModifiedDate = "";
            string LongName = "";
            string Status = "";
            string DomainHost = "";
            string DomainGUID = "";
            string DomainSID = "";
            string DomainControllerName = "";
            string IPAddress = "";

            // Create and structure the result DataTable
            DataTable dtDevices = new DataTable();
            dtDevices.Columns.Add("DeviceName");
            dtDevices.Columns.Add("LongName");
            dtDevices.Columns.Add("ObjectGUID");
            dtDevices.Columns.Add("ObjectSID");
            dtDevices.Columns.Add("CreatedDate");
            dtDevices.Columns.Add("ModifiedDate");
            dtDevices.Columns.Add("Status");
            dtDevices.Columns.Add("Domainname");
            dtDevices.Columns.Add("DomainGUID");
            dtDevices.Columns.Add("DomainSID");
            dtDevices.Columns.Add("DomainControllerName");
            dtDevices.Columns.Add("IPAddress");

            // Fetch domain information using helper method
            var domainInfo = GetDomainInfo();
            DomainHost = domainInfo.DomainHost;
            DomainGUID = domainInfo.DomainGUID;
            DomainSID = domainInfo.DomainSID;
            DomainControllerName = domainInfo.DomainControllerName;
            IPAddress = domainInfo.IPAddress;

            // If the DirectoryEntry is valid, search for Devices
            if (entry != null)
            {
                srch.Filter = "(&(objectClass=computer))"; // Searches only computer objects
                Thread.Sleep(1000);
                try
                {
                    SearchResultCollection srcColl = srch.FindAll();
                    if (srcColl != null)
                    {
                        // Loop through Device properties and add to DataTable
                        foreach (SearchResult srDevice in srcColl)
                        {
                            if (srDevice.Properties.Contains("cn"))
                            {
                                DeviceName = srDevice.Properties["cn"][0].ToString();
                            }
                            Guid deviceGUID = new Guid((byte[])srDevice.Properties["objectGUID"][0]);
                            byte[] sidBytes = (byte[])srDevice.Properties["objectSid"][0];
                            System.Security.Principal.SecurityIdentifier sid = new System.Security.Principal.SecurityIdentifier(sidBytes, 0);
                            ObjectSID = sid.Value;

                            if (srDevice.Properties.Contains("whenCreated"))
                            {
                                CreatedDate = srDevice.Properties["whenCreated"][0].ToString();
                            }
                            if (srDevice.Properties.Contains("whenChanged"))
                            {
                                ModifiedDate = srDevice.Properties["whenChanged"][0].ToString();
                            }

                            string dnParts = srDevice.Path.Split('/').Last();
                            LongName = objHelper.getOULong(srDevice.Path.Split(','));  // Assuming objHelper.getOULong handles path conversion

                            if (srDevice.Properties.Contains("userAccountControl"))
                            {
                                int userAccountControl = (int)srDevice.Properties["userAccountControl"][0];
                                bool isAccountDisabled = (userAccountControl & 0x2) != 0; // Check if the 'Account Disabled' bit is set

                                Status = isAccountDisabled ? "Disabled" : "Enabled";
                            }


                            DataRow dr = dtDevices.NewRow();
                            dr["DeviceName"] = DeviceName;
                            dr["LongName"] = dnParts;  
                            dr["ObjectGUID"] = deviceGUID.ToString(); 
                            dr["ObjectSID"] = ObjectSID;  
                            dr["CreatedDate"] = CreatedDate;
                            dr["ModifiedDate"] = ModifiedDate;
                            dr["Status"] = Status;
                            dr["Domainname"] = DomainHost;
                            dr["DomainGUID"] = DomainGUID;
                            dr["DomainSID"] = DomainSID;
                            dr["DomainControllerName"] = DomainControllerName;
                            dr["IPAddress"] = IPAddress;

                            dtDevices.Rows.Add(dr);// Add row to DataTable
                        }
                    }
                }
                catch (Exception ex)
                {
                    DataRow errorRow = dtDevices.NewRow();
                    errorRow["DeviceName"] = "Error";
                    errorRow["LongName"] = "N/A";
                    errorRow["ObjectGUID"] = "N/A";
                    errorRow["ObjectSID"] = "N/A";
                    errorRow["CreatedDate"] = "N/A";
                    errorRow["ModifiedDate"] = "N/A";
                    errorRow["Status"] = "N/A";
                    errorRow["Domainame"] = "Error fetching domain";
                    errorRow["DomainGUID"] = "N/A";
                    errorRow["DomainSID"] = "N/A";
                    dtDevices.Rows.Add(errorRow);
                }
            }
            return dtDevices; // Return the populated DataTable
        }

        /// <summary>
        /// Retrieves all Groups from Active Directory and returns their details in a DataTable.
        /// Also includes domain-related metadata such as Domain GUID, SID, and Domain Controller information.
        /// </summary>
        public DataTable GetAllGroups()
        {// Variable declarations for Group and domain metadata
            string GroupName = "";
            string Description = "";
            string ObjectGUID = "";
            string ObjectSID = "";
            string CreatedDate = "";
            string ModifiedDate = "";
            string LongName = "";
            string DomainHost = "";
            string DomainGUID = "";
            string DomainSID = "";
            string DomainControllerName = "";
            string IPAddress = "";

            // Create and structure the result DataTable
            DataTable dtGroups = new DataTable();
            dtGroups.Columns.Add("GroupName");
            dtGroups.Columns.Add("Description");
            dtGroups.Columns.Add("LongName");
            dtGroups.Columns.Add("ObjectGUID");
            dtGroups.Columns.Add("ObjectSID");
            dtGroups.Columns.Add("CreatedDate");
            dtGroups.Columns.Add("ModifiedDate");
            dtGroups.Columns.Add("Domainname");
            dtGroups.Columns.Add("DomainGUID");
            dtGroups.Columns.Add("DomainSID");
            dtGroups.Columns.Add("DomainControllerName");
            dtGroups.Columns.Add("IPAddress");
            // Fetch domain information using helper method
            var domainInfo = GetDomainInfo();
            DomainHost = domainInfo.DomainHost;
            DomainGUID = domainInfo.DomainGUID;
            DomainSID = domainInfo.DomainSID;
            DomainControllerName = domainInfo.DomainControllerName;
            IPAddress = domainInfo.IPAddress;

            // If the DirectoryEntry is valid, search for Grouops
            if (entry != null)
            {
                srch.Filter = "(&(objectClass=group))"; // fetch only group objects
                Thread.Sleep(1000);
                try
                {
                    objLoggerClass.Write("List", "GetAllGroups", 0,"Start GetAllGroups", "Trying to fetch groups..."); objLogWriter.WriteLogsFromQueue();
                    SearchResultCollection srcColl = srch.FindAll();
                    if (srcColl != null)
                    {
                        // Loop through Group properties and add to DataTable
                        foreach (SearchResult srGroup in srcColl)
                        {
                            objLoggerClass.Write("List", "GetAllGroups", 0, "Start GetAllGroups", "Processing Groups"+srGroup.Path); objLogWriter.WriteLogsFromQueue();
                            
                            if (srGroup.Properties.Contains("cn") && srGroup.Properties["cn"].Count > 0)
                            {
                                GroupName = srGroup.Properties["cn"][0].ToString();
                                objLoggerClass.Write("List", "GetAllGroups", 0, "Start GetAllGroups", "Group name" + GroupName); objLogWriter.WriteLogsFromQueue();
                            }
                            if (srGroup.Properties.Contains("description") && srGroup.Properties["description"].Count > 0)
                            {
                                Description = srGroup.Properties["description"][0].ToString();
                                objLoggerClass.Write("List", "GetAllGroups", 0, "Start GetAllGroups", "Description" + Description); objLogWriter.WriteLogsFromQueue();
                            }
                            if (srGroup.Properties.Contains("objectGUID") && srGroup.Properties["objectGUID"].Count > 0)
                            {
                                Guid groupGUID = new Guid((byte[])srGroup.Properties["objectGUID"][0]);
                                ObjectGUID = groupGUID.ToString();
                                objLoggerClass.Write("List", "GetAllGroups", 0, "Start GetAllGroups", "Group GUID" + ObjectGUID); objLogWriter.WriteLogsFromQueue();
                            }

                            if (srGroup.Properties.Contains("objectSid") && srGroup.Properties["objectSid"].Count > 0)
                            {
                                byte[] sidBytes = (byte[])srGroup.Properties["objectSid"][0];
                                System.Security.Principal.SecurityIdentifier sid = new System.Security.Principal.SecurityIdentifier(sidBytes, 0);
                                ObjectSID = sid.Value;
                                objLoggerClass.Write("List", "GetAllGroups", 0, "Start GetAllGroups", "Group SID" + ObjectSID); objLogWriter.WriteLogsFromQueue();
                            }

                            if (srGroup.Properties.Contains("whenCreated") && srGroup.Properties["whenCreated"].Count > 0)
                            {
                                CreatedDate = srGroup.Properties["whenCreated"][0].ToString();
                                objLoggerClass.Write("List", "GetAllGroups", 0, "Start GetAllGroups", "Created Date" + CreatedDate); objLogWriter.WriteLogsFromQueue();
                            }

                            if (srGroup.Properties.Contains("whenChanged") && srGroup.Properties["whenChanged"].Count > 0)
                            {
                                ModifiedDate = srGroup.Properties["whenChanged"][0].ToString();
                                objLoggerClass.Write("List", "GetAllGroups", 0, "Start GetAllGroups", "Modified Date" + ModifiedDate); objLogWriter.WriteLogsFromQueue();
                            }


                            string dnParts = srGroup.Path.Split('/').Last();
                            LongName = objHelper.getOULong(srGroup.Path.Split(','));

                            DataRow dr = dtGroups.NewRow();
                            dr["GroupName"] = GroupName;
                            dr["Description"] = Description;
                            dr["LongName"] = dnParts;  
                            dr["ObjectGUID"] = ObjectGUID;  
                            dr["ObjectSID"] = ObjectSID;  
                            dr["CreatedDate"] = CreatedDate;
                            dr["ModifiedDate"] = ModifiedDate;
                            dr["Domainname"] = DomainHost;
                            dr["DomainGUID"] = DomainGUID;
                            dr["DomainSID"] = DomainSID;
                            dr["DomainControllerName"] = DomainControllerName;
                            dr["IPAddress"] = IPAddress;
                            dtGroups.Rows.Add(dr); // Add row to DataTable
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("List", "groupList", 0, " get list of groups", "srcColl is null");
                        objLogWriter.WriteLogsFromQueue();
                    }
                }
                catch (Exception ex)
                {
                    objLoggerClass.Write("List", "groupList", 0, " get list of groups", "exception:" + ex);
                    objLogWriter.WriteLogsFromQueue();
                }
            }
            return dtGroups;// Return the populated DataTable
        }

        /// <summary>
        /// Retrieves all users from Active Directory along with the security groups (rights) they belong to.
        /// Each group a user is a member of is considered a 'right' and returned in a DataTable with User-Right pairs.
        /// </summary>
        public DataTable GetUserRights()
        {
            // Create and define columns for the result DataTable
            DataTable dtUserRights = new DataTable();
            dtUserRights.Columns.Add("User");
            dtUserRights.Columns.Add("Right");

            try
            {
                if (Connect())
                {
                    // Set LDAP filter to get only user objects
                    srch.Filter = "(objectClass=user)";
                    srch.PropertiesToLoad.Add("samaccountname");
                    srch.PropertiesToLoad.Add("memberOf");

                    SearchResultCollection results = srch.FindAll();

                    // Loop through each user object returned
                    foreach (SearchResult result in results)
                    {
                        string userName = result.Properties["samaccountname"][0].ToString();

                        if (result.Properties["memberOf"] != null)
                        {
                            foreach (var group in result.Properties["memberOf"])
                            {
                                DataRow dr = dtUserRights.NewRow();
                                dr["User"] = userName;
                                dr["Right"] = group.ToString(); // Group name represents the right assigned
                                dtUserRights.Rows.Add(dr);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
            return dtUserRights;  // Return the populated DataTable with user rights
        }

        /// <summary>
        /// Retrieves NTFS permissions (ACLs) for a given file or folder path.
        /// Returns a DataTable containing identity, access control type (Allow/Deny),
        /// rights (Read, Write, etc.), and whether the permission is inherited.
        /// </summary>
        public DataTable GetFileFolderPermissions(string path)
        {
            DataTable permissionsTable = new DataTable();
            permissionsTable.Columns.Add("Identity", typeof(string));// User or group name
            permissionsTable.Columns.Add("AccessControlType", typeof(string));// Allow or Deny
            permissionsTable.Columns.Add("Rights", typeof(string)); // Access rights (Read, Write, etc.)
            permissionsTable.Columns.Add("IsInherited", typeof(bool)); // True if permission is inherited

            // Check if the path is a file or directory
            if (Directory.Exists(path))
            {
                DirectorySecurity directorySecurity = Directory.GetAccessControl(path);
                objHelper.AddSecurityInfoToTable(directorySecurity, permissionsTable);
            }
            // Check if the given path is a file
            else if (System.IO.File.Exists(path))
            {
                FileSecurity fileSecurity = System.IO.File.GetAccessControl(path);
                objHelper.AddSecurityInfoToTable(fileSecurity, permissionsTable);
            }
            else
            {
                Console.WriteLine("The path does not exist.");
            }

            return permissionsTable;    // Return the populated DataTable
        }

        /// <summary>
        /// Retrieves all Group Policies from Active Directory and returns their details in a DataTable.
        /// Also includes domain-related metadata such as Domain GUID, SID, and Domain Controller information.
        /// </summary>
        public DataTable GetGroupPolicies()
        {
            // Create a DataTable to store group policies
            DataTable dtGroupPolicies = new DataTable();
            dtGroupPolicies.Columns.Add("Name");
            dtGroupPolicies.Columns.Add("DN");
            dtGroupPolicies.Columns.Add("CreationTime");
            dtGroupPolicies.Columns.Add("LastModifiedTime");
            dtGroupPolicies.Columns.Add("VersionNumber");
            dtGroupPolicies.Columns.Add("FileSysPath");
            dtGroupPolicies.Columns.Add("MachineExtensions");
            dtGroupPolicies.Columns.Add("UserExtensions");
            dtGroupPolicies.Columns.Add("GUID");
            dtGroupPolicies.Columns.Add("SID");
            dtGroupPolicies.Columns.Add("Status");
            dtGroupPolicies.Columns.Add("Domainname");
            dtGroupPolicies.Columns.Add("DomainGUID");
            dtGroupPolicies.Columns.Add("DomainSID");
            dtGroupPolicies.Columns.Add("DomainControllerName");
            dtGroupPolicies.Columns.Add("IPAddress");
            string DomainHost = "";
            string DomainGUID = "";
            string DomainSID = "";
            string DomainControllerName = "";
            string IPAddress = "";

            // Fetch domain information using helper method
            var domainInfo = GetDomainInfo();
            DomainHost = domainInfo.DomainHost;
            DomainGUID = domainInfo.DomainGUID;
            DomainSID = domainInfo.DomainSID;
            DomainControllerName = domainInfo.DomainControllerName;
            IPAddress = domainInfo.IPAddress;
            try
            {
                using (DirectoryEntry entry = new DirectoryEntry(string.Format("LDAP://{0}", DomainName), DomainUsername, DomainUserPassword))
                {
                    using (DirectorySearcher searcher = new DirectorySearcher(entry))
                    {// If the DirectoryEntry is valid, search for GPOs
                        searcher.Filter = "(objectClass=groupPolicyContainer)"; // filter only GPOs data 

                        searcher.PropertiesToLoad.Add("displayName");
                        searcher.PropertiesToLoad.Add("distinguishedName");
                        searcher.PropertiesToLoad.Add("whenCreated");
                        searcher.PropertiesToLoad.Add("whenChanged");
                        searcher.PropertiesToLoad.Add("versionNumber");
                        searcher.PropertiesToLoad.Add("gPCFileSysPath");
                        searcher.PropertiesToLoad.Add("gPCMachineExtensionNames");
                        searcher.PropertiesToLoad.Add("gPCUserExtensionNames");
                        searcher.PropertiesToLoad.Add("objectGuid");
                        searcher.PropertiesToLoad.Add("objectSid");
                        searcher.PropertiesToLoad.Add("flags");

                        SearchResultCollection results = searcher.FindAll();

                        // Loop through Group Policy properties and add to DataTable
                        foreach (SearchResult result in results)
                        {
                            string gpoName = result.Properties["displayName"].Count > 0 ? result.Properties["displayName"][0].ToString() : string.Empty;
                            string distinguishedName = result.Properties["distinguishedName"].Count > 0 ? result.Properties["distinguishedName"][0].ToString() : string.Empty;
                            string creationTime = result.Properties["whenCreated"].Count > 0 ? result.Properties["whenCreated"][0].ToString() : string.Empty;
                            string modifiedTime = result.Properties["whenChanged"].Count > 0 ? result.Properties["whenChanged"][0].ToString() : string.Empty;
                            string versionNumber = result.Properties["versionNumber"].Count > 0 ? result.Properties["versionNumber"][0].ToString() : string.Empty;
                            string fileSysPath = result.Properties["gPCFileSysPath"].Count > 0 ? result.Properties["gPCFileSysPath"][0].ToString() : string.Empty;
                            string machineExtensions = result.Properties["gPCMachineExtensionNames"].Count > 0 ? result.Properties["gPCMachineExtensionNames"][0].ToString() : string.Empty;
                            string userExtensions = result.Properties["gPCUserExtensionNames"].Count > 0 ? result.Properties["gPCUserExtensionNames"][0].ToString() : string.Empty;
                            string guid = result.Properties["objectGuid"].Count > 0
                                ? new Guid((byte[])result.Properties["objectGuid"][0]).ToString()
                                : string.Empty;

                            string status = result.Properties["flags"].Count > 0
                                ? DecodeGPOStatus(Convert.ToInt32(result.Properties["flags"][0]))
                                : "Unknown";

                            DataRow dr = dtGroupPolicies.NewRow();
                            dr["Name"] = gpoName;
                            dr["DN"] = distinguishedName;
                            dr["CreationTime"] = creationTime;
                            dr["LastModifiedTime"] = modifiedTime;
                            dr["GUID"] = guid;
                            dr["Status"] = status;
                            dr["Domainname"] = DomainHost;
                            dr["DomainGUID"] = DomainGUID;
                            dr["DomainSID"] = DomainSID;
                            dr["VersionNumber"] = versionNumber;
                            dr["FileSysPath"] = fileSysPath;
                            dr["MachineExtensions"] = machineExtensions;
                            dr["UserExtensions"] = userExtensions;
                            dr["DomainControllerName"] = DomainControllerName;
                            dr["IPAddress"] = IPAddress;
                            
                            dtGroupPolicies.Rows.Add(dr);// Add row to DataTable
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DataRow errorRow = dtGroupPolicies.NewRow();
                errorRow["Name"] = "Error";
                errorRow["DN"] = ex.Message;
                dtGroupPolicies.Rows.Add(errorRow);
            }
            return dtGroupPolicies;// Return the populated DataTable
        }
        public string DecodeGPOStatus(int flags)
        {
            switch (flags)
            {
                case 0: return "Enabled";
                case 1: return "User Configuration Disabled";
                case 2: return "Computer Configuration Disabled";
                case 3: return "All Disabled";
                default: return "Unknown";
            }
        }

        //private string MapExtensions(string extensions)
        //{
        //    // Use a regular expression to extract GUIDs from the input string
        //    var regex = new Regex(@"\{([0-9A-Fa-f-]+)\}"); // Match GUIDs enclosed in {}
        //    var matches = regex.Matches(extensions);

        //    // Use a StringBuilder for efficient concatenation
        //    StringBuilder result = new StringBuilder();

        //    foreach (Match match in matches)
        //    {
        //        if (match.Success)
        //        {
        //            string guid = match.Groups[1].Value; // Extract the GUID value
        //            string extensionName = GetPolicyExtensionName(guid);
        //            result.AppendLine($"{guid} - {extensionName}");
        //        }
        //    }

        //    return result.ToString().Trim();
        //}
        private string GetPolicyExtensionName(string guid)
    {
        try
        {
            string regPath = $@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\GPExtensions\{guid}";
            using (RegistryKey key = Registry.LocalMachine.OpenSubKey(regPath))
            {
                if (key != null)
                {
                    return key.GetValue("DisplayName", "Unknown Extension").ToString();
                }
            }
        }
        catch
        {
            return "Error retrieving name";
        }

        return "Unknown Extension";
    }



    }
}
