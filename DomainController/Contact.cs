using Logger;
using OwnYITCommon;
using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TectonaDatabaseHandlerDLL;

namespace DomainController
{
    public class Contact
    {
        public static string logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\AssertYIT\\DomainController\\";
        //string applicationName = "";
        LoggerClass objLoggerClass = new LoggerClass("");
        LogWriter objLogWriter = new LogWriter(logFilePath);
        helperMethod objHelper = new helperMethod();
        DatabaseCommon objCommon = new DatabaseCommon();
        DataTableConversion objDT = new DataTableConversion();
        DirectoryEntry entry;
        DirectorySearcher srch;
        bool isConnect = false;
        string DomainName = "";
        string DomainUsername = "";
        string DomainUserPassword = "";
        public Contact(string DomainName, string Username, string Password)
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
                objLoggerClass.Write("Contact", "Connect", 0, " can not connect to server ", "exception:" + ex);
                objLogWriter.WriteLogsFromQueue();
            }
            return isConnect;
        }
        public bool createContact(string Name, string surname, string mail, string phoneNumber)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        if (objHelper.IsContactExists(entry, Name))
                        {
                            DirectoryEntry newCont = entry.Children.Add($"CN={Name}", "contact");
                                newCont.Properties["displayName"].Value = Name;
                                newCont.Properties["sn"].Value = surname;
                                newCont.Properties["mail"].Value = mail;
                                newCont.Properties["telephoneNumber"].Value = phoneNumber;

                                newCont.CommitChanges();
                                bl= true;
                        }
                        else
                        {
                            objLoggerClass.Write("CreateContact", "ContactExists", 0, $"Contact '{Name}' already exists", "Contact creation failed due to duplicate"); objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("CreateContact", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("CreateContact", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("CreateContact", "createContact", 0, "create contact in DC", "Exception : " + ex);objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }


        public bool dltContact(string Name)
        {
            bool bl = false;
            try
            {
                if (Connect())
                {
                    if (objHelper.AreCredentialValid(entry))
                    {
                        DirectoryEntry contactEntry = entry.Children.Find($"CN={Name}", "contact");
                        if (contactEntry != null)
                        {
                            contactEntry.DeleteTree();
                            contactEntry.CommitChanges();
                            bl = true;
                        }
                        else
                        {
                            objLoggerClass.Write("deleteContact", "deleteContact", 0, "delete contact in DC", "contact not found");objLogWriter.WriteLogsFromQueue();
                        }
                    }
                    else
                    {
                        objLoggerClass.Write("deleteContact", "AccessDenied", 0, "Invalid domain credentials", "Credential validation failed"); objLogWriter.WriteLogsFromQueue();
                    }
                }
                else
                {
                    objLoggerClass.Write("deleteContact", "ConnectionError", 0, "Failed to connect to LDAP server", "Connection could not be established"); objLogWriter.WriteLogsFromQueue();
                }
            }
            catch (Exception ex)
            {
                objLoggerClass.Write("deleteContact", "deleteContact", 0, "delete contact in DC", "Exception : " + ex);objLogWriter.WriteLogsFromQueue();
            }
            return bl;
        }
    }
}
