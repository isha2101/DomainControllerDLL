using Logger;
using OwnYITCommon;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.DirectoryServices;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomainController
{
    public class AuditPolicy
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

        public AuditPolicy(string DomainName, string Username, string Password)
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

        //public DataTable GetAuditPolicies()
        //{
        //    DataTable auditPolicyTable = new DataTable();
        //    auditPolicyTable.Columns.Add("Category", typeof(string));
        //    auditPolicyTable.Columns.Add("Subcategory", typeof(string));
        //    auditPolicyTable.Columns.Add("Inclusions", typeof(string));

        //    try
        //    {
        //        ProcessStartInfo psi = new ProcessStartInfo()
        //        {
        //            FileName = "powershell.exe",
        //            Arguments = "-Command \"auditpol /get /category:*\"",
        //            RedirectStandardOutput = true,
        //            RedirectStandardError = true,
        //            UseShellExecute = false,
        //            CreateNoWindow = true
        //        };

        //        using (Process process = new Process() { StartInfo = psi })
        //        {
        //            process.Start();
        //            string output = process.StandardOutput.ReadToEnd();
        //            Console.WriteLine("PowerShell Output:\n" + output); // Debugging
        //            while (!process.StandardOutput.EndOfStream)
        //            {
        //                string line = process.StandardOutput.ReadLine();
        //                if (!string.IsNullOrWhiteSpace(line) && !line.StartsWith("Category") && !line.StartsWith("---"))
        //                {
        //                    string[] parts = line.Split(new[] { "\t" }, StringSplitOptions.RemoveEmptyEntries);

        //                    //string[] parts = line.Split(new[] { "  " }, StringSplitOptions.RemoveEmptyEntries);
        //                    if (parts.Length >= 3)
        //                    {
        //                        auditPolicyTable.Rows.Add(parts[0].Trim(), parts[1].Trim(), parts[2].Trim());
        //                    }
        //                }
        //            }
        //            process.WaitForExit();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Error retrieving audit policies: {ex.Message}");
        //    }

        //    return auditPolicyTable;
        //}


        public DataTable GetAuditPolicies()
        {
            DataTable auditPolicyTable = new DataTable();
            auditPolicyTable.Columns.Add("Category", typeof(string));
            auditPolicyTable.Columns.Add("Subcategory", typeof(string));
            auditPolicyTable.Columns.Add("Inclusions", typeof(string));

            try
            {
                ProcessStartInfo psi = new ProcessStartInfo()
                {
                    FileName = "powershell.exe",
                    Arguments = "-Command \"auditpol /get /category:*\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process process = new Process() { StartInfo = psi })
                {
                    process.Start();
                    string output = process.StandardOutput.ReadToEnd();
                    Console.WriteLine("PowerShell Output:\n" + output); // Debugging

                    string currentCategory = string.Empty;
                    string[] lines = output.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

                    foreach (string line in lines)
                    {
                        if (!string.IsNullOrWhiteSpace(line))
                        {
                            if (!line.StartsWith("  ")) // Category line (No leading spaces)
                            {
                                currentCategory = line.Trim();
                            }
                            else // Subcategory line (Indented)
                            {
                                string[] parts = line.Trim().Split(new[] { "  " }, StringSplitOptions.RemoveEmptyEntries);
                                if (parts.Length >= 2)
                                {
                                    auditPolicyTable.Rows.Add(currentCategory, parts[0].Trim(), parts[1].Trim());
                                }
                            }
                        }
                    }

                    process.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving audit policies: {ex.Message}");
            }

            return auditPolicyTable;
            
        }
         public bool enableAuditPolicy()
         {
            bool bl = false;
            DataTable auditPolicies = GetAuditPolicies();
            foreach(DataRow dr in  auditPolicies.Rows)
            {
                string category = dr["Categoy"].ToString();
                string subCategory = dr["Subcategory"].ToString();
                string status = dr["Inclusions"].ToString();

                if(status.Equals("No Auditing", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"Enabling auditing for: {subCategory}");
                    string command = $"-Command \"auditpol /set /subcategory:'{subCategory}' /success:enable /failure:enable\"";
                    ProcessStartInfo psi = new ProcessStartInfo()
                    {
                        FileName = "powershell.exe",
                        Arguments = command,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    };
                    using (Process process = new Process() { StartInfo = psi })
                    {
                        process.Start();
                        string output = process.StandardOutput.ReadToEnd();
                        string error = process.StandardError.ReadToEnd();
                        process.WaitForExit();

                        if (!string.IsNullOrWhiteSpace(error))
                        {
                            Console.WriteLine($"Error setting {subCategory}: {error}");
                        }
                        else
                        {
                            Console.WriteLine($"Successfully enabled auditing for: {subCategory}");
                            bl = true;
                        }
                    }
                }
            }
            return bl;
         }

    }
}
