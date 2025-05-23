using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DomainController
{
    public class TCDomianController
    {
        public void testing()
        {
            //AuditPolicy objAudit = new AuditPolicy("192.168.100.60", "DEMO\\Administrator", "RESEt21");
            //objAudit.Connect();
            //DataTable auditPolicyStatus =  objAudit.GetAuditPolicies();
            //bool enableAuditPolicy = objAudit.enableAuditPolicy();

            //User objUser = new User("192.168.100.60", "DEMO\\Administrator", "RESEt21");
            //objUser.Connect();
            //bool create = objUser.CreateUser("het", "hetP@demo.lab", "hetP", "het from .net team", "PhetP@43210");
            //bool update = objUser.updateUser("rajesh", "Update User", "Updated user from C# team");
            //bool resetPwd = objUser.resetPassword("rajesh", "Reset@12345");
            //bool enable = objUser.enableUser("rajesh");
            //bool disable = objUser.disableUser("rajesh");
            //bool move = objUser.moveUserToOU("rajesh", "OU=newOU");
            //bool delete = objUser.deleteUser("het");
            //string Status = objUser.getUserPwdStatus("jeel");
            //DataTable dt = objUser.getPwdPolicy();
            //objUser.setCustomPwd("jeel");


            //Device objDevice = new Device("192.168.100.60", "DEMO\\Administrator", "RESEt21");
            //objDevice.Connect();
            //bool deviceJoinDomain = objDevice.IsMachineJoinedToDomain();
            //bool createDevice = objDevice.CreateDevice("newDeviceI", "new device from c# team");
            //bool updateDevice = objDevice.updateDevice("newDeviceI", "Updated device");
            //bool enableDevice = objDevice.enableDevice("newDeviceI");
            //bool disableDevice = objDevice.disableDevice("newDeviceI");
            //bool moveDevice = objDevice.moveDeviceToOU("newDeviceI", "OU=isha_newOU");
            //bool dltDevice = objDevice.deleteDevice("newDeviceI");

            //Group objGroup = new Group("192.168.100.60", "DEMO\\Administrator", "RESEt21");
            ////objGroup.Connect();
            //bool createGrp = objGroup.createGroup("iGroup", "new group from C#");
            //DataTable grpInfo = objGroup.grpInformation("iGroup");
            //bool addUser = objGroup.addUserToGrp("rajput", "iGroup");
            //bool dltUser = objGroup.dltUserFromGrp("rajput", "iGroup");
            //bool moveGrp = objGroup.moveGrpToOU("iGroup", "OU=isha_newOU");
            //bool dltGroup = objGroup.deleteGroup("iGroup");

            //OrganizationalUnit objOU = new OrganizationalUnit("192.168.100.60", "DEMO\\Administrator", "RESEt21");
            //objOU.Connect();
            //bool createOU = objOU.createOU("iOrgUnit", "OU from C# team");
            //bool movOU = objOU.moveOU("newOU", "isha_newOU");
            //bool modifyOU = objOU.modifyOU("isha_newOU", "description", "Successfully updated OU");
            //bool dltOU = objOU.dltOU("iOrgUnit");

            //GroupPolicy objGP = new GroupPolicy("192.168.100.60", "DEMO\\Administrator", "RESEt21");
            // objGP.Connect();
            // DataTable policyData = objGP.GetAuditPolicies();
            //bool crtConfigGPO = objGP.createAndConfigureGPO("new GPO", "Demo.lab");
            //bool crtGPO = objGP.createGPO("newGPO", "Demo.lab");
            //bool linkGPO = objGP.linkGPOToOU("newGPO", "newOU", "Demo.lab");
            //bool unlinkGPO = objGP.unlinkGPOFromOU("newGPO", "newOU", "Demo.lab");
            //bool dltGPO = objGP.deleteGPO("newGPO", "Demo.lab");

            //Linkage objLink = new Linkage("192.168.100.249", "tec\\Administrator", "TectonA$123!@#");
            //objLink.Connect();

            //DataTable OUGrpLinkage = objLink.GetOUGroupLinkage();
            //DataTable GrpUserLinkage = objLink.GetGroupUserLinkage();
            //DataTable GrpDeviceLinkage = objLink.GetGroupDeviceLinkage();
            ////DataTable deviceUserLinkage = objLink.GetDeviceUserLinkage("shukla");
            //DataTable GetGPODeviceLinkage = objLink.GetGPODeviceLinkage();
            //DataTable GetGPOsUserLinakge = objLink.GetGPOUserLinkage();
            //DataTable GPOsOULinkage = objLink.GetGPOsOULinkage();
            //DataTable ouUserLinkage = objLink.GetOUUserLinkage();
            //DataTable ouDeviceLinkage = objLink.GetOUDeviceLinkage();

            List objList = new List("192.168.100.115", "tectonas\\Administrator", "Tectona#123");
            objList.Connect();
            //DataTable getGPOList = objList.GetGrpPolicies();
            //DataTable getServiceAcc = objList.GetServiceAccountList();
            //DataTable GetHighPrivilegeUser = objList.GetHighPrivilegeUser();
            //DataTable serviceAcc = objList.GetUserPwd();
            DataTable strOUList = objList.GetAllOU();
            //DataTable strUser = objList.GetAllUsers();
            //DataTable strComputer = objList.GetAllDevices();
            //DataTable strGroups = objList.GetAllGroups();
            //DataTable gpoList = objList.GetGroupPolicies();
            //DataTable userRights = objList.GetUserRights();
            //DataTable filefolderList = objList.GetFileFolderPermissions("C:\\DOTNET\\PROJECTS\\");


            //Contact objContact = new Contact("192.168.100.60", "DEMO\\Administrator", "RESEt21");
            //objContact.Connect();
            //bool crtcontact = objContact.createContact("newcontact", "desai", "icon123@gmail.com", "1234567890");
            //bool dltcontact = objContact.dltContact("newcontact");
        }
    }
}
