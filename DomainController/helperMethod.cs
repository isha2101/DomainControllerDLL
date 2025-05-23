using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.Linq;
using System.Management.Automation;
using System.Text;
using System.Threading.Tasks;
using System.Security.Principal;
using System.Data;
using System.Security.AccessControl;
using System.Runtime.InteropServices;

namespace DomainController
{
    public class helperMethod
    {
        public bool AreCredentialValid(DirectoryEntry entry)
        {
            try
            {
                object nativeObject = entry.NativeObject;
                return true;
            }
            catch
            {
                return false;
            }
        }
        public bool SetPassword(DirectoryEntry newUser, string password)
        {
            try
            {
                //newUser.Invoke("setpassword", new object[] { password });
                newUser.Invoke("SetPassword", new object[] { password });
                return true;
            }
            catch 
            { return false; }
        }
        public bool isUserExists(DirectoryEntry entry ,string AccountName)
        {
            using(DirectorySearcher searcher = new DirectorySearcher(entry) )
            {
                searcher.Filter = $"(sAMAccountName={AccountName})";
                searcher.PropertiesToLoad.Add("sAMAccountName");
                SearchResult result = searcher.FindOne();
                return result == null;
            }
        }

        public bool isDeviceExists(DirectoryEntry entry ,string DeviceName)
        {
            using(DirectorySearcher searcher  = new DirectorySearcher(entry) )
            {
                searcher.Filter = $"(sAMAccountName={DeviceName})";
                searcher.PropertiesToLoad.Add("sAMAccountName");
                SearchResult result = searcher.FindOne();
                return result == null;
            }
        }
        public bool isGrpExists(DirectoryEntry entry ,string GroupName)
        {
            using(DirectorySearcher searcher = new DirectorySearcher (entry) )
            {
                searcher.Filter = $"(&(objectClass=group)(sAMAccountName={GroupName}))";
                searcher.PropertiesToLoad.Add("sAMAccountName");
                SearchResult result = searcher.FindOne();
                return result == null;
            }
        }
        public bool IsOUExists(DirectoryEntry entry, string ouName)
        {
            using (DirectorySearcher searcher = new DirectorySearcher(entry))
            {
                searcher.Filter = $"(&(objectClass=organizationalUnit)(ou={ouName}))";
                searcher.PropertiesToLoad.Add("ou");
                SearchResult result = searcher.FindOne();
                return result == null;
            }
        }

        public bool IsContactExists(DirectoryEntry entry, string contactName)
        {
            using (DirectorySearcher searcher = new DirectorySearcher(entry))
            {
                searcher.Filter = $"(&(objectClass=contact)(cn={contactName}))";
                searcher.PropertiesToLoad.Add("cn");
                SearchResult result = searcher.FindOne();
                return result == null;
            }
        }

        public int ConvertLargeIntegerToDays(object largeInteger)
        {
            if (largeInteger == null)
            {
                return 0;
            }
            Type type = largeInteger.GetType();
            int highPart = (int)type.InvokeMember("HighPart", System.Reflection.BindingFlags.GetProperty, null, largeInteger, null);
            int lowPart = (int)type.InvokeMember("LowPart", System.Reflection.BindingFlags.GetProperty, null, largeInteger, null);
            long ticks = ((long)highPart << 32) + (uint)lowPart;
            TimeSpan timeSpan = TimeSpan.FromTicks(ticks);

            return (int)Math.Abs(timeSpan.TotalDays); // Use absolute value for positive days
        }

        public long ConvertLargeIntegerToLong(object largeInteger)
        {
            // The IADsLargeInteger interface exposes HighPart and LowPart properties
            Type type = largeInteger.GetType();
            int highPart = (int)type.InvokeMember("HighPart", System.Reflection.BindingFlags.GetProperty, null, largeInteger, null);
            int lowPart = (int)type.InvokeMember("LowPart", System.Reflection.BindingFlags.GetProperty, null, largeInteger, null);
            return ((long)highPart << 32) + (uint)lowPart;
        }
        public void ValidateUserPassword(string userPath, string userName, string password)
        {
            try
            {
                using (DirectoryEntry validationEntry = new DirectoryEntry(userPath, userName, password))
                {
                    // Attempt to bind to the directory
                    object nativeObject = validationEntry.NativeObject;
                    Console.WriteLine("Password reset successfully and validated.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Password reset but validation failed: " + ex.Message);
            }
        }
        public string getOULong(string[] strOU)
        {
            string str = "";
            for (int i = 0; i < strOU.Length; i++)
            {
                if (strOU[i].StartsWith("OU"))
                    str = " >> " + strOU[i].Split(new string[] { "OU=" }, StringSplitOptions.None)[1] + str;
                else if (strOU[i].Contains("OU"))
                    str = " >> " + strOU[i].Split(new string[] { "OU=" }, StringSplitOptions.None)[1] + str;
            }
            if (str.Length > 0)
                str = str.Substring(4);
            return str;
        }
   
        public string getOULongName(string ouDN)
        {
            string[] Parts = ouDN.Split(',');
            string longName = "";

            for (int i = Parts.Length - 1; i >= 0; i--)
            {
                if (Parts[i].StartsWith("OU=") || Parts[i].StartsWith("CN=") || Parts[i].StartsWith("cn="))
                {
                    longName = longName + " >> " + Parts[i].Substring(3); // Remove "OU=" or "CN="
                }
            }

            return longName.TrimStart('>', ' '); // Remove leading ">>" if exists
        }

        public string get_User_PC_Long(string[] strUser)
        {
            string str = "";
            for (int i = 0; i < strUser.Length; i++)
            {
                if (strUser[i].StartsWith("CN"))
                    str = " >> " + strUser[i].Split(new string[] { "CN=" }, StringSplitOptions.None)[1] + str;
                else if (strUser[i].Contains("CN"))
                    str = " >> " + strUser[i].Split(new string[] { "CN=" }, StringSplitOptions.None)[1] + str;
            }
            if (str.Length > 4)
                str = str.Substring(4);
            return str;
        }
        public bool CheckForErrors(PowerShell ps, string operation)
        {
            if (ps.Streams.Error.Count > 0)
            {
                foreach (var error in ps.Streams.Error)
                {
                    Console.WriteLine($"Error {operation}: " + error);
                }
                return false;
            }
            return true;
        }
        public string GetUserNameFromSID(string sid)
        {
            try
            {
                SecurityIdentifier sidObj = new SecurityIdentifier(sid);
                return sidObj.Translate(typeof(NTAccount)).ToString();
            }
            catch (Exception ex)
            {
                return "UnknownUser";
            }
        }

        public  void AddSecurityInfoToTable(FileSystemSecurity security, DataTable table)
        {
            AuthorizationRuleCollection rules = security.GetAccessRules(true, true, typeof(System.Security.Principal.NTAccount));

            foreach (FileSystemAccessRule rule in rules)
            {
                // Create a new row and populate it with the rule details
                DataRow row = table.NewRow();
                row["Identity"] = rule.IdentityReference.Value;
                row["AccessControlType"] = rule.AccessControlType.ToString();
                row["Rights"] = rule.FileSystemRights.ToString();
                row["IsInherited"] = rule.IsInherited;

                // Add the row to the table
                table.Rows.Add(row);
            }
        }

        public string ConvertSidToUsername(string sid)
        {
            try
            {
                SecurityIdentifier sidObject = new SecurityIdentifier(sid);
                NTAccount account = (NTAccount)sidObject.Translate(typeof(NTAccount));
                return account.ToString();
            }
            catch (Exception ex)
            {
                return $"Error converting SID to username: {ex.Message}";
            }
        }

        //public long ConvertLargeIntegerToLong(object largeIntegerObj)
        //{
        //    if (largeIntegerObj is IADsLargeInteger largeInteger)
        //    {
        //        return ((long)largeInteger.HighPart << 32) + largeInteger.LowPart;
        //    }
        //    return 0;
        //}

        //[ComImport, Guid("9068270B-0939-11D1-8BE1-00C04FD8D503")]
        //[InterfaceType(ComInterfaceType.InterfaceIsDual)]
        //public interface IADsLargeInteger
        //{
        //    int HighPart { get; set; }
        //    int LowPart { get; set; }
        //}
    }

}

