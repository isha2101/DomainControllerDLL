
using Logger;
using Newtonsoft.Json.Converters;
using OwnYITCommon;
using System;
using System.Collections.Generic;
using System.Data;
using System.DirectoryServices;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Web.UI.WebControls.WebParts;
using static System.Windows.Forms.LinkLabel;

namespace DomainController
{
    public class GroupPolicy
    {
        public static string logFilePath = Environment.GetFolderPath(Environment.SpecialFolder.CommonApplicationData) + "\\AssertYIT\\DomainController\\";
        LoggerClass objLoggerClass = new LoggerClass("");
        LogWriter objLogWriter = new LogWriter(logFilePath);
        helperMethod objHelper = new helperMethod();
        DataTableConversion objdatc = new DataTableConversion();
        DirectoryEntry entry;
        DirectorySearcher srch;

        string DomainName = "";
        string DomainUsername = "";
        string DomainUserPassword = "";

        public GroupPolicy(string DomainName, string Username, string Password)
        {
            this.DomainName = DomainName;
            this.DomainUsername = Username;
            this.DomainUserPassword = Password;
        }
        public bool Connect()
        {
            bool bl = false;
            try
            {
                entry = new DirectoryEntry(string.Format("LDAP://{0}", DomainName), DomainUsername, DomainUserPassword);
                srch = new DirectorySearcher(entry);
                bl = true;
            }
            catch (Exception ex)
            {
            }
            return bl;
        }

        public DataTable GetAuditPolicies()
        {
            // Create a DataTable to store audit policies
            DataTable auditPolicyTable = new DataTable();
            auditPolicyTable.Columns.Add("Category", typeof(string));
            auditPolicyTable.Columns.Add("SubCategory", typeof(string));
            auditPolicyTable.Columns.Add("Status", typeof(string));

            try
            {
                // Run the "auditpol" command
                Process process = new Process();
                process.StartInfo.FileName = "auditpol";
                process.StartInfo.Arguments = "/get /category:*";
                process.StartInfo.RedirectStandardOutput = true;
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.CreateNoWindow = true;

                process.Start();

                // Read output from the command
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();

                // Parse the output and populate the DataTable
                 string[] lines = output.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

                //foreach (string line in lines)
                //{
                //    Console.WriteLine($"Line: {line}");
                //    // Check if line contains at least 3 columns
                //    if (line.Trim().Length > 0)
                //    {
                //        // Split using regular expression
                //        string[] parts = regex.Split(line.Trim());
                //        Console.WriteLine($"Part: {parts}");
                //        if (parts.Length == 3)
                //        {
                //            string category = parts[0].Trim();
                //            string subCategory = parts[1].Trim();
                //            string setting = parts[2].Trim();

                //            // Add row to the DataTable
                //            auditPolicyTable.Rows.Add(category, subCategory, setting);
                //        }
                //    }
                //}
                
                string currentCategory = string.Empty;

                foreach (string line in lines)
                {
                    // Trim whitespace from the line
                    string trimmedLine = line.Trim();

                    // Skip headers and empty lines
                    if (string.IsNullOrWhiteSpace(trimmedLine) || trimmedLine.StartsWith("Category/Subcategory") || trimmedLine.Equals("System audit policy"))
                    {
                        continue;
                    }

                    // If the line has no leading spaces, it is a Category
                    if (!line.StartsWith(" "))
                    {
                        currentCategory = trimmedLine; // Update the current category
                    }
                    else
                    {
                        // This is a SubCategory and Setting
                        // Use regex to split into SubCategory and Setting
                        Regex regex = new Regex(@"\s{2,}"); // Matches 2 or more spaces
                        string[] parts = regex.Split(trimmedLine);

                        if (parts.Length == 2) // Ensure the line has exactly two parts
                        {
                            string subCategory = parts[0].Trim();
                            string setting = parts[1].Trim();

                            // Add row to the DataTable
                            auditPolicyTable.Rows.Add(currentCategory, subCategory, setting);
                        }
                    }
                }



            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }

            // Return the populated DataTable
            return auditPolicyTable;
        }


        
        public bool createGPO(string gpoName, string domainName)
        {
            try
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    // Set the execution policy to RemoteSigned for the current process
                    ps.AddCommand("Set-ExecutionPolicy")
                      .AddArgument("RemoteSigned")
                      .AddParameter("Scope", "Process");

                    ps.Invoke();
                    if (!objHelper.CheckForErrors(ps, "creating GPO"))
                    {
                        return false;
                    }
                    // Clear the commands and proceed to load the GroupPolicy module
                    ps.Commands.Clear();

                    string modulePath = @"C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules\GroupPolicy\GroupPolicy.psd1";

                    ps.AddCommand("Import-Module")
                      .AddArgument(modulePath);

                    ps.Invoke();

                    if (!objHelper.CheckForErrors(ps, "creating GPO"))
                    {
                        return false;
                    }
                    // Add the command to create the GPO
                    ps.Commands.Clear();
                    ps.AddCommand("New-GPO")
                      .AddParameter("Name", gpoName)
                      .AddParameter("Domain", domainName);

                    var result = ps.Invoke();
                 
                    if (!objHelper.CheckForErrors(ps, "creating GPO"))
                    {
                        return false;
                    }
                    ps.Commands.Clear();
                    return true;
                }
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public bool deleteGPO(string gpoName, string domainName)
        {
            try
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    // Set the execution policy to RemoteSigned for the current process
                    ps.AddCommand("Set-ExecutionPolicy")
                      .AddArgument("RemoteSigned")
                      .AddParameter("Scope", "Process");

                    ps.Invoke();

                    //if (ps.Streams.Error.Count > 0)
                    //{
                    //    foreach (var error in ps.Streams.Error)
                    //    {
                    //        Console.WriteLine("Error setting execution policy: " + error);
                    //    }
                    //    return false;
                    //}
                    if (!objHelper.CheckForErrors(ps, "deleting GPO"))
                    {
                        return false;
                    }
                    // Clear the commands and load the GroupPolicy module
                    ps.Commands.Clear();

                    string modulePath = @"C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules\GroupPolicy\GroupPolicy.psd1";

                    ps.AddCommand("Import-Module")
                      .AddArgument(modulePath);

                    ps.Invoke();

                    // Check for errors while importing the module
                    //if (ps.Streams.Error.Count > 0)
                    //{
                    //    foreach (var error in ps.Streams.Error)
                    //    {
                    //        Console.WriteLine("Error loading module: " + error);
                    //    }
                    //    return false;
                    //}
                    if (!objHelper.CheckForErrors(ps, "deleting GPO"))
                    {
                        return false;
                    }
                    // Clear previous commands and add the command to delete the GPO
                    ps.Commands.Clear();
                    ps.AddCommand("Remove-GPO")
                      .AddParameter("Name", gpoName)
                      .AddParameter("Domain", domainName);

                    ps.Invoke();

                    // Check for errors during deletion
                    //if (ps.Streams.Error.Count > 0)
                    //{
                    //    foreach (var error in ps.Streams.Error)
                    //    {
                    //        Console.WriteLine("Error deleting GPO: " + error);
                    //    }
                    //    return false;
                    //}
                    if (!objHelper.CheckForErrors(ps, "deleting GPO"))
                    {
                        return false;
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return false;
            }
        }


        public bool linkGPOToOU(string gpoName, string ouName, string domainName)
        {
            try
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    // Set the execution policy to RemoteSigned for the current process
                    ps.AddCommand("Set-ExecutionPolicy")
                      .AddArgument("RemoteSigned")
                      .AddParameter("Scope", "Process");

                    ps.Invoke();

                    //if (ps.Streams.Error.Count > 0)
                    //{
                    //    foreach (var error in ps.Streams.Error)
                    //    {
                    //        Console.WriteLine("Error setting execution policy: " + error);
                    //    }
                    //    return false;
                    //}
                    if (!objHelper.CheckForErrors(ps, "link GPO"))
                    {
                        return false;
                    }
                    // Clear the commands and load the GroupPolicy module
                    ps.Commands.Clear();

                    string modulePath = @"C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules\GroupPolicy\GroupPolicy.psd1";

                    ps.AddCommand("Import-Module")
                      .AddArgument(modulePath);

                    ps.Invoke();

                    // Check for errors while importing the module
                    //if (ps.Streams.Error.Count > 0)
                    //{
                    //    foreach (var error in ps.Streams.Error)
                    //    {
                    //        Console.WriteLine("Error loading module: " + error);
                    //    }
                    //    return false;
                    //}
                    if (!objHelper.CheckForErrors(ps, "link GPO"))
                    {
                        return false;
                    }
                    // Add the command to link the GPO to the OU
                    ps.Commands.Clear();
                    ps.AddCommand("New-GPLink")
                      .AddParameter("Name", gpoName)  // GPO name
                      .AddParameter("Target", $"OU={ouName},DC={domainName.Split('.')[0]},DC={domainName.Split('.')[1]}");  // OU and domain name in distinguished name (DN) format

                    var result = ps.Invoke();

                    // Check for errors while linking GPO
                    //if (ps.Streams.Error.Count > 0)
                    //{
                    //    foreach (var error in ps.Streams.Error)
                    //    {
                    //        Console.WriteLine("Error linking GPO: " + error);
                    //    }
                    //    return false;
                    //}
                    if (!objHelper.CheckForErrors(ps, "link GPO"))
                    {
                        return false;
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return false;
            }
        }

        public bool unlinkGPOFromOU(string gpoName, string ouName, string domainName)
        {
            try
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    // Set the execution policy to RemoteSigned for the current process
                    ps.AddCommand("Set-ExecutionPolicy")
                      .AddArgument("RemoteSigned")
                      .AddParameter("Scope", "Process");

                    ps.Invoke();

                    if (!objHelper.CheckForErrors(ps, "unlink GPO"))
                    {
                        return false;
                    }

                    // Clear the commands and load the GroupPolicy module
                    ps.Commands.Clear();

                    string modulePath = @"C:\WINDOWS\system32\WindowsPowerShell\v1.0\Modules\GroupPolicy\GroupPolicy.psd1";

                    ps.AddCommand("Import-Module")
                      .AddArgument(modulePath);

                    ps.Invoke();

                    if (!objHelper.CheckForErrors(ps, "unlink GPO"))
                    {
                        return false;
                    }

                    // Add the command to unlink the GPO from the OU
                    ps.Commands.Clear();
                    ps.AddCommand("Remove-GPLink")
                      .AddParameter("Name", gpoName)  // GPO name
                      .AddParameter("Target", $"OU={ouName},DC={domainName.Split('.')[0]},DC={domainName.Split('.')[1]}")  // OU and domain name in distinguished name (DN) format
                      .AddParameter("Confirm", false);  // Avoid confirmation prompt

                    var result = ps.Invoke();

                    if (!objHelper.CheckForErrors(ps, "unlink GPO"))
                    {
                        return false;
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return false;
            }
        }

        public bool createAndConfigureGPO(string gpoName, string domainName)
        {
            try
            {
                using (PowerShell ps = PowerShell.Create())
                {
                    // Set the execution policy
                    ps.AddCommand("Set-ExecutionPolicy")
                      .AddArgument("RemoteSigned")
                      .AddParameter("Scope", "Process");
                    ps.Invoke();
                    if (!objHelper.CheckForErrors(ps, "setting execution policy"))
                    {
                        return false;
                    }

                    // Import GroupPolicy module
                    ps.Commands.Clear();
                    ps.AddCommand("Import-Module")
                      .AddArgument("GroupPolicy");
                    ps.Invoke();
                    if (!objHelper.CheckForErrors(ps, "importing GroupPolicy module"))
                    {
                        return false;
                    }

                    // Create the GPO
                    ps.Commands.Clear();
                    ps.AddCommand("New-GPO")
                      .AddParameter("Name", gpoName)
                      .AddParameter("Domain", domainName);
                    ps.Invoke();
                    if (!objHelper.CheckForErrors(ps, "creating GPO"))
                    {
                        return false;
                    }

                    // Configure a setting (example: Hide the "All Apps" button from Start Menu)
                    ps.Commands.Clear();
                    ps.AddCommand("Set-GPRegistryValue")
                      .AddParameter("Name", gpoName)
                      .AddParameter("Key", @"Software\Microsoft\Windows\CurrentVersion\Policies\Explorer")
                      .AddParameter("ValueName", "NoStartMenuMorePrograms")
                      .AddParameter("Type", "DWord")
                      .AddParameter("Value", 1); // 1 = Enabled
                    ps.Invoke();
                    if (!objHelper.CheckForErrors(ps, "configuring GPO settings"))
                    {
                        return false;
                    }

                    return true;
                }
            }
            catch (Exception ex)
            {
                // Log exception or handle it
                return false;
            }
        }


    }
}
