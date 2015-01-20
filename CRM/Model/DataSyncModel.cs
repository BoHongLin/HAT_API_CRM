using CRM.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRM.Model
{
    public class DataSyncModel
    {
        private Guid datasyncID;
        public void CreateDataSyncForCRM(String str)
        {
            try
            {
                DateTime now = DateTime.Now;
                Entity datasync = new Entity("datasync_datasync");
                datasync["datasync_name"] = str + " - " + now.ToString();
                datasync["datasync_start"] = now;
                datasyncID = EnvironmentSetting.Service.Create(datasync);
                EnvironmentSetting.ErrorType = ErrorType.None;
            }
            catch (Exception ex)
            {
                EnvironmentSetting.ErrorMsg = "建立 DataSync 失敗\n";
                EnvironmentSetting.ErrorMsg += ex.Message + "\n";
                EnvironmentSetting.ErrorMsg += ex.Source + "\n";
                EnvironmentSetting.ErrorMsg += ex.StackTrace + "\n";
                EnvironmentSetting.ErrorType = ErrorType.DATASYNC;
            }
        }
        public void UpdateDataSyncForCRM(int succss, int fail, int partially)
        {
            try
            {
                Entity datasync = new Entity("datasync_datasync");
                datasync["datasync_datasyncid"] = datasyncID;
                datasync["datasync_end"] = DateTime.Now;
                datasync["datasync_success"] = new Decimal(succss);
                datasync["datasync_fail"] = new Decimal(fail);
                datasync["datasync_partially"] = new Decimal(partially);
                EnvironmentSetting.Service.Update(datasync);
                EnvironmentSetting.ErrorType = ErrorType.None;
            }
            catch (Exception ex)
            {
                EnvironmentSetting.ErrorMsg = "更新 DataSync 失敗\n";
                EnvironmentSetting.ErrorMsg += ex.Message + "\n";
                EnvironmentSetting.ErrorMsg += ex.Source + "\n";
                EnvironmentSetting.ErrorMsg += ex.StackTrace + "\n";
                EnvironmentSetting.ErrorType = ErrorType.DATASYNC;
            }
        }
        public void UpdateDataSyncWithErrorForCRM(String msg)
        {
            try
            {
                Entity datasync = new Entity("datasync_datasync");
                datasync["datasync_datasyncid"] = datasyncID;
                datasync["datasync_end"] = DateTime.Now;
                datasync["datasync_errormsg"] = msg;
                EnvironmentSetting.Service.Update(datasync);
                //EnvironmentSetting.ErrorType = ErrorType.None;
            }
            catch (Exception ex)
            {
                EnvironmentSetting.ErrorMsg = "更新 DataSyncErrorMsg 失敗\n";
                EnvironmentSetting.ErrorMsg += ex.Message + "\n";
                EnvironmentSetting.ErrorMsg += ex.Source + "\n";
                EnvironmentSetting.ErrorMsg += ex.StackTrace + "\n";
                EnvironmentSetting.ErrorType = ErrorType.DATASYNC;
            }
        }
        public void CreateDataSyncDetailForCRM(String primarykey,String name, TransactionType transactiontype, TransactionStatus transactionstatus)
        {
            try
            {
                Entity datasyncdetail = new Entity("datasync_datasyncdetail");
                datasyncdetail["datasync_datasyncid"] = new EntityReference("datasync_datasync", datasyncID);
                datasyncdetail["datasync_primarykey"] = primarykey;
                datasyncdetail["datasync_name"] = name;
                datasyncdetail["datasync_transactiontype"] = new OptionSetValue(Convert.ToInt32(transactiontype));
                datasyncdetail["datasync_transactionstatus"] = new OptionSetValue(Convert.ToInt32(transactionstatus));
                datasyncdetail["datasync_desc"] = EnvironmentSetting.ErrorMsg;
                EnvironmentSetting.Service.Create(datasyncdetail);
                EnvironmentSetting.ErrorType = ErrorType.None;
            }
            catch (Exception ex)
            {
                EnvironmentSetting.ErrorMsg = "建立 DataSyncDetail 失敗\n";
                EnvironmentSetting.ErrorMsg += ex.Message + "\n";
                EnvironmentSetting.ErrorMsg += ex.Source + "\n";
                EnvironmentSetting.ErrorMsg += ex.StackTrace + "\n";
                EnvironmentSetting.ErrorType = ErrorType.DATASYNCDETAIL;
            }
        }
    }
}
