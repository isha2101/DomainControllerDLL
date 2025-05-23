using Logger;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using OwnYITCommon;
using System;
using System.Collections.Generic;
using System.Data;
using System.DirectoryServices;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;

namespace DomainController
{
    public class Linkage
    {
        public static string logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\AssertYIT\\DomainController\\";
        LoggerClass objLoggerClass = new LoggerClass("");
        LogWriter objLogWriter = new LogWriter(logFilePath);
        DataTableConversion objDT = new DataTableConversion();
        helperMethod objHelper = new helperMethod();
        DirectoryEntry entry;
        DirectorySearcher srch;
        bool isConnect = false;
        string DomainName = "";
        string DomainUsername = "";
        string DomainUserPassword = "";

        public Linkage(string DomainName, string Username, string Password)
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
                objLoggerClass.Write("Linkage", "Connect", 0, " can not connect to server ", "exception:" + ex);
                objLogWriter.WriteLogsFromQueue();
            }
            return isConnect;
        }

        /// <summary>
        /// Retrieves a mapping between Group Policy Objects (GPOs) and the devices (computers) they are linked to,
        /// This method queries all Organizational Units (OUs) in Active Directory, fetches GPO links from each OU,
        /// then finds all computers in those OUs and associates them with the linked GPOs.
        /// </summary>
        public DataTable GetGPODeviceLinkage()
        {
            DataTable dtGPOUserLinkage = new DataTable();
            dtGPOUserLinkage.Columns.Add("GPO");
            dtGPOUserLinkage.Columns.Add("Device");
            dtGPOUserLinkage.Columns.Add("IPAddress");
            try
            {
                // Apply filter to search only Organizational Units (OUs)
                if (entry != null)
                {
                    srch.Filter = "(&(objectClass=organizationalUnit))";
                };

                SearchResultCollection ouResults = srch.FindAll();

                if (ouResults != null)
                {
                    foreach (SearchResult srOU in ouResults)
                    {
                        DirectoryEntry ouEntry = new DirectoryEntry(srOU.Path, DomainUsername, DomainUserPassword);

                        // Get GPO links from the OU's gPLink attribute
                        object gpoLinks = ouEntry.Properties["gPLink"].Value; 

                        if (gpoLinks != null)
                        {
                            // Extract GPO names
                            string[] gpoLinksArray = gpoLinks.ToString().Split(new string[] { "[LDAP://cn=", "[LDAP://CN=" }, StringSplitOptions.None).Skip(1).Select(g => "CN=" + g.Split(';')[0]).ToArray();

                            // Search for all computer objects under this OU
                            DirectorySearcher userSearcher = new DirectorySearcher(ouEntry)
                            {
                                Filter = "(&(objectClass=computer))"
                            };
                            SearchResultCollection userResults = userSearcher.FindAll();
                            string IPAddress = $"{DomainName}";
                            foreach (SearchResult srUser in userResults)
                            {
                                string deviceName = "";
                                ResultPropertyCollection userProps = srUser.Properties;

                                // Retrieve the computer name
                                if (userProps["cn"] != null && userProps["cn"].Count > 0)
                                {
                                    deviceName = userProps["cn"][0].ToString();
                                }
                                // Map each GPO to the device
                                foreach (string gpoName in gpoLinksArray)
                                {
                                    DataRow drLinkage = dtGPOUserLinkage.NewRow();
                                    drLinkage["GPO"] = gpoName;
                                    drLinkage["Device"] = deviceName;
                                    drLinkage["IPAddress"] = IPAddress;
                                    dtGPOUserLinkage.Rows.Add(drLinkage);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DataRow errorRow = dtGPOUserLinkage.NewRow();
                errorRow["GPO"] = "Error";
                errorRow["Device"] = ex.Message;

                dtGPOUserLinkage.Rows.Add(errorRow);
            }
            return dtGPOUserLinkage;
        }

        /// <summary>
        /// Retrieves a mapping between Group Policy Objects (GPOs) and the users they are linked to,
        /// This method searches all Organizational Units (OUs) in Active Directory, reads GPO links from each OU,
        /// finds all user accounts within those OUs, and then associates each user with the linked GPOs.
        /// </summary>
        public DataTable GetGPOUserLinkage()
        {
            DataTable dtGPOUserLinkage = new DataTable();
            dtGPOUserLinkage.Columns.Add("GPO");
            dtGPOUserLinkage.Columns.Add("User");
            dtGPOUserLinkage.Columns.Add("IPAddress");
            try
            {
                // Apply filter to get all Organizational Units (OUs) from Active Directory
                if (entry != null)
                {
                    srch.Filter = "(&(objectClass=organizationalUnit))";
                };

                SearchResultCollection ouResults = srch.FindAll();

                if (ouResults != null)
                {
                    foreach (SearchResult srOU in ouResults)
                    {
                        DirectoryEntry ouEntry = new DirectoryEntry(srOU.Path, DomainUsername, DomainUserPassword);

                        // Retrieve the gPLink property which holds GPOs linked to this OU
                        object gpoLinks = ouEntry.Properties["gPLink"].Value;

                        if (gpoLinks != null)
                        {
                            // Extract GPO distinguished names 
                            string[] gpoLinksArray = gpoLinks.ToString().Split(new string[] { "[LDAP://cn=", "[LDAP://CN=" }, StringSplitOptions.None).Skip(1).Select(g => "CN=" + g.Split(';')[0]).ToArray();

                             // Search for user accounts under this OU
                            DirectorySearcher userSearcher = new DirectorySearcher(ouEntry)
                            {
                                Filter = "(&(objectCategory=User)(objectClass=person))"
                            };
                            SearchResultCollection userResults = userSearcher.FindAll();
                            string IPAddress = $"{DomainName}";
                            foreach (SearchResult srUser in userResults)
                            {
                                string userName = "";
                                ResultPropertyCollection userProps = srUser.Properties;

                                // Get the user's common name (CN)
                                if (userProps["cn"] != null && userProps["cn"].Count > 0)
                                {
                                    userName = userProps["cn"][0].ToString();
                                }
                                // Map each linked GPO to the user
                                foreach (string gpoName in gpoLinksArray)
                                {
                                    DataRow drLinkage = dtGPOUserLinkage.NewRow();
                                    drLinkage["GPO"] = gpoName;
                                    drLinkage["User"] = userName;
                                    drLinkage["IPAddress"] = IPAddress;
                                    dtGPOUserLinkage.Rows.Add(drLinkage);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DataRow errorRow = dtGPOUserLinkage.NewRow();
                errorRow["GPO"] = "Error";
                errorRow["User"] = ex.Message;
                dtGPOUserLinkage.Rows.Add(errorRow);
            }
            return dtGPOUserLinkage;
        }

        /// <summary>
        /// Retrieves a mapping between Group Policy Objects (GPOs) and the Organizational Units (OUs) they are linked to,
        /// This function searches all Organizational Units (OUs) in Active Directory, extracts linked GPOs from the gPLink property
        /// of each OU, and maps each GPO to its corresponding OU.
        /// </summary>
        public DataTable GetGPOsOULinkage()
        {
            // Create a DataTable to store the GPO and OU linkage
            DataTable dtGPOOULinkage = new DataTable();
            dtGPOOULinkage.Columns.Add("GPO");
            dtGPOOULinkage.Columns.Add("OU");
            dtGPOOULinkage.Columns.Add("IPAddress");
            try
            {
                // Set search filter to retrieve only Organizational Units (OUs)
                if (entry != null)
                {
                    srch.Filter = "(&(objectClass=organizationalUnit))";
                };

                SearchResultCollection ouResults = srch.FindAll();

                if (ouResults != null)
                {
                    foreach (SearchResult srOU in ouResults)
                    {
                        string ouDN = srOU.Path.Split('/').Last(); 
                        //string ouLongName = objHelper.getOULong(srOU.Path.Split(',')); // Fully qualified OU name (if needed)
                        string ouLongName = objHelper.getOULongName(ouDN);

                        DirectoryEntry ouEntry = new DirectoryEntry(srOU.Path, DomainUsername, DomainUserPassword);
                        string IPAddress = $"{DomainName}";

                        // Get the gPLink property of the OU
                        object gpoLinks = ouEntry.Properties["gPLink"].Value;

                        if (gpoLinks != null)
                        {
                            string[] gpoLinksArray = gpoLinks.ToString()
                                    .Split(new string[] { "[LDAP://cn=", "[LDAP://CN=" }, StringSplitOptions.None)
                                    .Skip(1) // Skip the initial empty segment
                                    .Select(g => "CN=" + g.Split(';')[0]) // Extract the required part and prepend "CN="
                                    .ToArray();
                            foreach (string gpoName in gpoLinksArray)
                            {
                                // Add GPO-OU mapping to the DataTable
                                DataRow drLinkage = dtGPOOULinkage.NewRow();
                                drLinkage["GPO"] = gpoName;
                                drLinkage["OU"] = ouLongName;
                                drLinkage["IPAddress"] = IPAddress;
                                dtGPOOULinkage.Rows.Add(drLinkage);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error or handle it according to your needs
                DataRow errorRow = dtGPOOULinkage.NewRow();
                errorRow["GPO"] = "Error";
                errorRow["OU"] = ex.Message;
                dtGPOOULinkage.Rows.Add(errorRow);
            }

            // Return the DataTable containing the GPO to OU mapping
            return dtGPOOULinkage;
        }

        /// <summary>
        /// Retrieves a mapping between users and their deepest Organizational Unit (OU) in Active Directory.
        /// The method iterates through all OUs, finds users within each, and determines the most specific OU (deepest in the hierarchy)
        /// each user belongs to. It returns this information in a DataTable containing the OU name, user name, and IP/domain name.
        /// </summary>
        public DataTable GetOUUserLinkage()
        {
            // Create a DataTable to store the OU and User linkage
            DataTable dtProperty = new DataTable();
            dtProperty.Columns.Add("OU");
            dtProperty.Columns.Add("User");
            dtProperty.Columns.Add("IPAddress");
            try
            {
                if (entry != null)
                {
                    // Set the filter to retrieve only Organizational Units
                    srch.Filter = "(&(objectClass=organizationalUnit))";
                    SearchResultCollection ouResults = srch.FindAll();

                    if (ouResults != null)
                    {
                        // Dictionary to store the deepest OU for each user
                        Dictionary<string, string> userToOUMap = new Dictionary<string, string>();

                        foreach (SearchResult srOU in ouResults)
                        {
                            string ouDN = srOU.Path.Split('/').Last();
                            string ouLongName = objHelper.getOULong(srOU.Path.Split(','));
                            //string ouLongName = objHelper.getOULongName(ouDN);
                            using (DirectoryEntry ouEntry = new DirectoryEntry(srOU.Path, DomainUsername, DomainUserPassword))
                            {
                                DirectorySearcher userSearcher = new DirectorySearcher(ouEntry);
                                userSearcher.Filter = "(&(objectCategory=User)(objectClass=person))";
                                SearchResultCollection userResults = userSearcher.FindAll();
                                
                                foreach (SearchResult srUser in userResults)
                                {
                                    string userName = "";
                                    ResultPropertyCollection userProps = srUser.Properties;

                                    // Retrieve the user name (samaccountname)
                                    foreach (string userPropName in userProps.PropertyNames)
                                    {
                                        if (userPropName == "cn")
                                        {
                                            userName = userProps[userPropName][0].ToString();
                                        }
                                    }
                                    // If user already exists in dictionary, keep the deepest OU (based on hierarchy depth)
                                    if (userToOUMap.ContainsKey(userName))
                                    {
                                        if (userToOUMap[userName].Split(new string[] { ">>" }, StringSplitOptions.None).Length < ouLongName.Split(new string[] { ">>" }, StringSplitOptions.None).Length)
                                        {
                                            userToOUMap[userName] = ouLongName;
                                            
                                        }
                                    }
                                    else
                                    {
                                        userToOUMap[userName] = ouLongName;
                                    }
                                }
                            }
                        }

                        // Add the final user-OU mapping to the DataTable
                        foreach (var entry in userToOUMap)
                        {
                            DataRow drLinkage = dtProperty.NewRow();
                            drLinkage["OU"] = entry.Value;
                            drLinkage["User"] = entry.Key;
                          
                            dtProperty.Rows.Add(drLinkage);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error or handle it according to your needs

                DataRow errorRow = dtProperty.NewRow();
                errorRow["OU"] = "Error";
                errorRow["User"] = ex.Message;
                dtProperty.Rows.Add(errorRow);
            }

            // Return the DataTable containing the OU to User mapping
            return dtProperty;
        }

        /// <summary>
        /// Retrieves a mapping between computers (devices) and their deepest Organizational Unit (OU) in Active Directory.
        /// The method iterates through all OUs, finds computers within each, and determines the most specific OU
        /// (deepest in hierarchy) each device belongs to. The result is returned as a DataTable with columns for OU and Device.
        /// </summary>
        public DataTable GetOUDeviceLinkage()
        {
            string strOUDeviceLinkage = "";
            string strPropertiesJSON = "";
            DataTable dtProperty = new DataTable();
            dtProperty.Columns.Add("OU");
            dtProperty.Columns.Add("Device");

            try
            {
                if (entry != null)
                {
                    // Filter to get only Organizational Units
                    srch.Filter = "(&(objectClass=organizationalUnit))";
                    SearchResultCollection ouResults = srch.FindAll();

                    if (ouResults != null)
                    {
                        // Dictionary to store the deepest OU for each device
                        Dictionary<string, string> deviceToOUMap = new Dictionary<string, string>();

                        foreach (SearchResult srOU in ouResults)
                        {
                            string ouDN = srOU.Path.Split('/').Last();
                            string ouLongName = objHelper.getOULong(srOU.Path.Split(','));
                            //string ouLongName = objHelper.getOULongName(ouDN);

                            using (DirectoryEntry ouEntry = new DirectoryEntry(srOU.Path, DomainUsername, DomainUserPassword))
                            { 
                                // Search for computer objects within this OU
                                DirectorySearcher deviceSearcher = new DirectorySearcher(ouEntry);
                                deviceSearcher.Filter = "(&(objectCategory=computer)(objectClass=computer))";
                                SearchResultCollection deviceResults = deviceSearcher.FindAll();

                                foreach (SearchResult srDevice in deviceResults)
                                {
                                    string deviceName = "";
                                    DataTable dtDeviceProperties = new DataTable();
                                    dtDeviceProperties.Columns.Add("Property");
                                    dtDeviceProperties.Columns.Add("Value");

                                    ResultPropertyCollection deviceProps = srDevice.Properties;

                                    // Extract device properties and build a property table
                                    foreach (string devicePropName in deviceProps.PropertyNames)
                                    {
                                        DataRow dr = dtDeviceProperties.NewRow();
                                        string devicePropValue = deviceProps[devicePropName][0].ToString();
                                        dr["Property"] = devicePropName;
                                        dr["Value"] = devicePropValue;
                                        dtDeviceProperties.Rows.Add(dr);

                                        if (devicePropName == "name")
                                        {
                                            deviceName = devicePropValue;
                                        }
                                    }
                                    strPropertiesJSON = objDT.DataTableToJSONString(dtDeviceProperties);

                                    // Determine if this is the deepest OU seen for this device
                                    if (deviceToOUMap.ContainsKey(deviceName))
                                    {
                                        if (deviceToOUMap[deviceName].Split(new string[] { " >> " }, StringSplitOptions.None).Length < ouLongName.Split(new string[] { " >> " }, StringSplitOptions.None).Length)
                                        {
                                            deviceToOUMap[deviceName] = ouLongName;
                                        }
                                    }
                                    else
                                    {
                                        deviceToOUMap[deviceName] = ouLongName;
                                    }
                                   
                                }
                            }
                        }

                        // Add the deepest OU for each device to the DataTable
                        foreach (var entry in deviceToOUMap)
                        {
                            DataRow drLinkage = dtProperty.NewRow();
                            drLinkage["OU"] = entry.Value;
                            drLinkage["Device"] = entry.Key;
                            dtProperty.Rows.Add(drLinkage);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("linkage", "OUDeviceLinkage", 0, " get linkage of OU and Device", "exception:" + ex);
                objLogWriter.WriteLogsFromQueue();
            }
            return dtProperty;
        }

        /// <summary>
        /// Retrieves a mapping between Active Directory Organizational Units (OUs) and security groups contained within them.
        /// For each OU found in the directory, the method searches for group objects and adds their names along with the OU name
        /// and domain IP address to a DataTable. This mapping helps understand which groups are located in which OUs.
        /// </summary>
        public DataTable GetOUGroupLinkage()
        {
            // Create a DataTable to store the OU and Group linkage
            DataTable dtProperty = new DataTable();
            dtProperty.Columns.Add("OU");
            dtProperty.Columns.Add("Group");
            dtProperty.Columns.Add("IPAddress");
            try
            {
                if (entry != null)
                {
                    // Search filter to get all Organizational Units
                    srch.Filter = "(&(objectClass=organizationalUnit))";
                    SearchResultCollection ouResults = srch.FindAll();

                    if (ouResults != null)
                    {
                        foreach (SearchResult srOU in ouResults)
                        {
                            string ouDN = srOU.Path.Split('/').Last();
                            string ouLongName = objHelper.getOULong(srOU.Path.Split(','));
                            //string ouLongName = objHelper.getOULongName(ouDN);

                            using (DirectoryEntry ouEntry = new DirectoryEntry(srOU.Path, DomainUsername, DomainUserPassword))
                            {
                                // Search for group objects inside this OU
                                DirectorySearcher groupSearcher = new DirectorySearcher(ouEntry);
                                groupSearcher.Filter = "(&(objectCategory=Group)(objectClass=group))";
                                SearchResultCollection groupResults = groupSearcher.FindAll();
                                string IPAddress = $"{DomainName}";
                                if (groupResults.Count > 0)
                                {

                                    foreach (SearchResult srGroup in groupResults)
                                    {
                                        string groupName = "";
                                        ResultPropertyCollection groupProps = srGroup.Properties;

                                        // Retrieve the group name (samaccountname or cn)
                                        foreach (string groupPropName in groupProps.PropertyNames)
                                        {
                                            if (groupPropName == "samaccountname") // or use "cn" if preferred
                                            {
                                                groupName = groupProps[groupPropName][0].ToString();
                                                break;
                                            }
                                        }
                                        // If a group name is found, add to DataTable
                                        if (!string.IsNullOrEmpty(groupName))
                                        {
                                            DataRow drLinkage = dtProperty.NewRow();
                                            drLinkage["OU"] = ouLongName;
                                            drLinkage["Group"] = groupName;
                                            drLinkage["IPAddress"] = IPAddress;
                                            dtProperty.Rows.Add(drLinkage);
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
                // Log error or handle it according to your needs

                DataRow errorRow = dtProperty.NewRow();
                errorRow["OU"] = "Error";
                errorRow["Group"] = ex.Message;
                dtProperty.Rows.Add(errorRow);
            }

            // Return the DataTable containing the OU to Group mapping
            return dtProperty;
        }


        /// <summary>
        /// Retrieves a mapping between Active Directory groups and the computer (device) accounts that are members of those groups.
        /// For each group found in the directory, it identifies members that are computers and logs their names along with the group name
        /// and domain information into a DataTable. This helps in auditing or understanding which devices belong to which security groups.
        /// </summary>
        public DataTable GetGroupDeviceLinkage()
        {
            // Create a DataTable to store the Group and Device linkage
            DataTable dtProperty = new DataTable();
            dtProperty.Columns.Add("Group");
            dtProperty.Columns.Add("Device");
            dtProperty.Columns.Add("IPAddress");

            try
            {
                if (entry != null)
                {
                    // Search filter to find all groups in Active Directory
                    srch.Filter = "(&(objectClass=group))";
                    SearchResultCollection groupResults = srch.FindAll();

                    if (groupResults != null)
                    {
                        foreach (SearchResult srGroup in groupResults)
                        {
                            string groupName = "";
                            ResultPropertyCollection groupProps = srGroup.Properties;

                            // Retrieve the group name (samaccountname or name)
                            foreach (string groupPropName in groupProps.PropertyNames)
                            {
                                if (groupPropName == "samaccountname")
                                {
                                    groupName = groupProps[groupPropName][0].ToString();
                                }
                            }

                            // Retrieve the members of the group
                            if (groupProps["member"] != null)
                            {
                                foreach (var member in groupProps["member"])
                                {
                                    string deviceDN = member.ToString();

                                    try
                                    {
                                        using (DirectoryEntry deviceEntry = new DirectoryEntry($"LDAP://{DomainName}/{deviceDN}", DomainUsername, DomainUserPassword))
                                        
                                        {
                                            DirectorySearcher deviceSearcher = new DirectorySearcher(deviceEntry);
                                            deviceSearcher.Filter = "(&(objectCategory=computer)(objectClass=computer))";
                                            SearchResult deviceResult = deviceSearcher.FindOne();
                                            string IPAddress = $"{DomainName}";

                                            if (deviceResult != null)
                                            {
                                                string deviceName = "";
                                                ResultPropertyCollection deviceProps = deviceResult.Properties;

                                                foreach (string devicePropName in deviceProps.PropertyNames)
                                                {
                                                    if (devicePropName == "name")
                                                    {
                                                        deviceName = deviceProps[devicePropName][0].ToString();
                                                    }
                                                }

                                                // Add the Group-Device linkage to the DataTable
                                                DataRow drLinkage = dtProperty.NewRow();
                                                drLinkage["Group"] = groupName;
                                                drLinkage["Device"] = deviceName;
                                                drLinkage["IPAddress"] = IPAddress;
                                                dtProperty.Rows.Add(drLinkage);
                                            }
                                        }
                                    }
                                    catch (Exception deviceEx)
                                    {
                                        // Handle errors while processing individual devices
                                        DataRow errorRow = dtProperty.NewRow();
                                        errorRow["Group"] = groupName;
                                        errorRow["Device"] = deviceEx.Message;
                                        dtProperty.Rows.Add(errorRow);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("linkage", "GroupDeviceLinkage", 0, " get linkage of Group and Device", "exception:" + ex);
                objLogWriter.WriteLogsFromQueue();
            }

            // Return the DataTable containing the Group to Device mapping
            return dtProperty;
        }

        /// <summary>
        /// Retrieves a mapping between Active Directory groups and the user accounts that are members of those groups.
        /// For each group found in the directory, it identifies members that are users and logs their names along with the group name
        /// and domain information into a DataTable. This helps in auditing or understanding which userss belong to which security groups.
        /// </summary>
        public DataTable GetGroupUserLinkage()
        {
            // Create a DataTable to store the Group and User linkage
            DataTable dtProperty = new DataTable();
            dtProperty.Columns.Add("Group");
            dtProperty.Columns.Add("User");
            dtProperty.Columns.Add("IPAddress");
            try
            {
                if (entry != null)
                {
                    // search for filter only groups
                    srch.Filter = "(&(objectClass=group))";
                    SearchResultCollection groupResults = srch.FindAll();

                    if (groupResults != null)
                    {
                        foreach (SearchResult srGroup in groupResults)
                        {
                            string groupName = "";
                            ResultPropertyCollection groupProps = srGroup.Properties;

                            // Retrieve the group name (samaccountname or name)
                            foreach (string groupPropName in groupProps.PropertyNames)
                            {
                                if (groupPropName == "samaccountname")
                                {
                                    groupName = groupProps[groupPropName][0].ToString();
                                }
                            }

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
                                            string IPAddress = $"{DomainName}";

                                            if (userResult != null)
                                            {
                                                string userName = "";
                                                ResultPropertyCollection userProps = userResult.Properties;

                                                foreach (string userPropName in userProps.PropertyNames)
                                                {
                                                    if (userPropName == "cn")
                                                    {
                                                        userName = userProps[userPropName][0].ToString();
                                                    }
                                                }

                                                // Add the Group-User linkage to the DataTable
                                                DataRow drLinkage = dtProperty.NewRow();
                                                drLinkage["Group"] = groupName;
                                                drLinkage["User"] = userName;
                                                drLinkage["IPAddress"] = IPAddress;
                                                dtProperty.Rows.Add(drLinkage);
                                            }
                                            else
                                            {

                                            }
                                        }
                                    }
                                    catch (Exception userEx)
                                    {
                                        // Handle errors while processing individual users
                                        DataRow errorRow = dtProperty.NewRow();
                                        errorRow["Group"] = groupName;
                                        errorRow["User"] = userEx.Message;
                                        dtProperty.Rows.Add(errorRow);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error or handle it according to your needs
                DataRow errorRow = dtProperty.NewRow();
                errorRow["Group"] = "Error";
                errorRow["User"] = ex.Message;
                dtProperty.Rows.Add(errorRow);
            }

            // Return the DataTable containing the Group to User mapping
            return dtProperty;
        }





    }
}
