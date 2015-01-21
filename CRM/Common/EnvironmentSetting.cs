using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.InteropServices;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;

namespace CRM.Common
{
    public class EnvironmentSetting
    {

        [DllImport("kernel32")]
        private static extern long WritePrivateProfileString(string section, string key, string val, string filePath);
        [DllImport("kernel32")]
        private static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);

        #region
        private static String errorMsg = "";
        private static ErrorType errorType = ErrorType.None;
        private static StringBuilder DB_SERVER = new StringBuilder(255);  //回傳所要接收的值
        private static StringBuilder DB_USERNAME = new StringBuilder(255);
        private static StringBuilder DB_PASSWORD = new StringBuilder(255);
        private static StringBuilder CRM_SERVER = new StringBuilder(255);
        private static StringBuilder CRM_DOMAIN = new StringBuilder(255);
        private static StringBuilder CRM_USERNAME = new StringBuilder(255);
        private static StringBuilder CRM_PASSWORD = new StringBuilder(255);
        private static StringBuilder CRM_UOMSCHEDULE = new StringBuilder(255);
        private static StringBuilder CRM_UOM = new StringBuilder(255);
        private static IOrganizationService service;
        private static OrganizationServiceContext xrm;
        private static SqlDataReader reader;
        private static SqlConnection con;
        #endregion

        #region
        public static StringBuilder USERNAME
        {
            get { return CRM_USERNAME; }
            set { CRM_USERNAME = value; }
        }
        public static IOrganizationService Service
        {
            get { return service; }
            set { service = value; }
        }
        public static OrganizationServiceContext Xrm
        {
            get { return xrm; }
            set { xrm = value; }
        }
        public static SqlDataReader Reader
        {
            get { return reader; }
            set { reader = value; }
        }
        public static String ErrorMsg
        {
            get { return errorMsg; }
            set { errorMsg = value; }
        }
        public static ErrorType ErrorType
        {
            get { return errorType; }
            set { errorType = value; }
        }
        public static StringBuilder UomScheduleName
        {
            get { return CRM_UOMSCHEDULE; }
            set { CRM_UOMSCHEDULE = value; }
        }
        public static StringBuilder Uom
        {
            get { return CRM_UOM; }
            set { CRM_UOM = value; }
        }
        #endregion

        public static void LoadSetting()
        {
            LoadINI();
            if (errorType == ErrorType.None)
            {
                LoadCRM();
            }
        }
        private static void LoadINI()
        {
            string path = Environment.CurrentDirectory + "\\setting.ini";
            if (System.IO.File.Exists(path) == false)
            {
                WritePrivateProfileString("DB_INFO", "SERVER", "192.168.0.1", path);
                WritePrivateProfileString("DB_INFO", "USERNAME", "username", path);
                WritePrivateProfileString("DB_INFO", "PASSWORD", "password", path);
                WritePrivateProfileString("CRM_INFO", "SERVER", "http://crm2011/crm", path);
                WritePrivateProfileString("CRM_INFO", "DOMAIN", "domain", path);
                WritePrivateProfileString("CRM_INFO", "USERNAME", "username", path);
                WritePrivateProfileString("CRM_INFO", "PASSWORD", "password", path);
                WritePrivateProfileString("CRM_INFO", "UOMSCHEDULE", "name", path);
                WritePrivateProfileString("CRM_INFO", "UOM", "name", path);
                errorMsg = "設定檔遺失...\n";
                errorMsg += "自動建立預設檔案完成，請進入資料夾內修改參數。\n";
                errorMsg += "檔案位置 : " + path + "\n";
                Console.WriteLine(errorMsg);
                errorType = ErrorType.INI;
            }
            else
            {
                try
                {
                    GetPrivateProfileString("DB_INFO", "SERVER", "null", DB_SERVER, 255, path);
                    GetPrivateProfileString("DB_INFO", "USERNAME", "null", DB_USERNAME, 255, path);
                    GetPrivateProfileString("DB_INFO", "PASSWORD", "null", DB_PASSWORD, 255, path);

                    GetPrivateProfileString("CRM_INFO", "SERVER", "null", CRM_SERVER, 255, path);
                    GetPrivateProfileString("CRM_INFO", "DOMAIN", "null", CRM_DOMAIN, 255, path);
                    GetPrivateProfileString("CRM_INFO", "USERNAME", "null", CRM_USERNAME, 255, path);
                    GetPrivateProfileString("CRM_INFO", "PASSWORD", "null", CRM_PASSWORD, 255, path);
                    GetPrivateProfileString("CRM_INFO", "UOMSCHEDULE", "null", CRM_UOMSCHEDULE, 255, path);
                    GetPrivateProfileString("CRM_INFO", "UOM", "null", CRM_UOM, 255, path);
                    errorType = ErrorType.None;
                }
                catch (Exception ex)
                {
                    errorMsg = "設定檔讀取失敗\n";
                    errorMsg += ex.Message + "\n";
                    Console.WriteLine(errorMsg);
                    errorType = ErrorType.INI;
                }
            }
        }
        private static void LoadCRM()
        {
            try
            {
                Uri organizationUri = new Uri(CRM_SERVER.ToString() + "/XRMServices/2011/Organization.svc");
                var cred = new ClientCredentials();
                cred.UserName.UserName = CRM_DOMAIN.ToString() + "\\" + CRM_USERNAME.ToString();
                cred.UserName.Password = CRM_PASSWORD.ToString();
                var serviceproxy = new OrganizationServiceProxy(organizationUri, null, cred, null);
                service = (IOrganizationService)serviceproxy;
                xrm = new OrganizationServiceContext(service);
            }
            catch (Exception ex)
            {
                errorMsg = "CRM連線失敗\n";
                errorMsg += ex.Message + "\n";
                Console.WriteLine(errorMsg);
                errorType =  ErrorType.CRM;
            }
        }
        public static void LoadDB()
        {
            try
            {
                con = new SqlConnection("Data Source=" + DB_SERVER + ";Initial Catalog=SynCRM;Persist Security Info=True;User ID=" + DB_USERNAME + ";Password=" + DB_PASSWORD);
                con.Open();
            }
            catch (Exception ex)
            {
                errorMsg = "ERP連線失敗\n";
                errorMsg += ex.Message + "\n";
                Console.WriteLine(errorMsg);
                errorType = ErrorType.DB;
            }
        }
        public static void GetList(String command)
        {
            try
            {
                reader = new SqlCommand(command, con).ExecuteReader();
            }
            catch (Exception ex)
            {
                errorMsg = "ERP資料讀取失敗\n";
                errorMsg += ex.Message + "\n";
                Console.WriteLine(errorMsg);
                errorType =  ErrorType.DB;
            }
        }
    }
}
