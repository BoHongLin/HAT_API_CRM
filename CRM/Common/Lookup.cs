using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CRM.Common
{
    public class Lookup
    {
        //List<Entity> lookuplist;
        //public Lookup(OrganizationServiceContext xrm, string realationEntityName)
        //{
        //    try
        //    {
        //        lookuplist = (from x in xrm.CreateQuery(realationEntityName) select x).ToList();
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("撈取資料失敗 : " + realationEntityName);
        //        Console.WriteLine(ex.Message);
        //    }
        //}

        //public Guid GetGuid(String value, String fieldName, String IDName)
        //{

        //    Entity entity = lookuplist.Find(x => x[fieldName].ToString() == value);
        //    if (entity == null)
        //        return Guid.Empty;
        //    else
        //        return new Guid(entity[IDName].ToString());


        //    //Guid defaultGuid = Guid.Empty;
        //    //foreach (var c in lookuplist)
        //    //{
        //    //    if (c[fieldName].ToString() == value)
        //    //    {
        //    //        Console.WriteLine(c[IDName].ToString());
        //    //        Console.ReadLine();
        //    //        return new Guid(c[IDName].ToString());
        //    //    }
        //    //}

        //    ////for (int i = 0, size = lookuplist.Count; i < size; i++)
        //    ////{
        //    ////    if (lookuplist[i][fieldName].ToString() == value)
        //    ////        return new Guid(lookuplist[i][IDName].ToString());
        //    ////}
        //    //return defaultGuid;
        //}

        public static Guid RetrieveEntityGuid(String EntityName, String value, String fieldName)
        {
            try
            {
                FilterExpression filter = new FilterExpression();

                //productnumber Equal pno
                ConditionExpression condition1 = new ConditionExpression();
                condition1.AttributeName = fieldName;
                condition1.Operator = ConditionOperator.Equal;
                condition1.Values.Add(value);
                filter.Conditions.Add(condition1);


                //EntityLogicalName
                QueryExpression qe = new QueryExpression(EntityName);
                qe.Criteria.AddFilter(filter);
                EntityCollection collections = EnvironmentSetting.Service.RetrieveMultiple(qe);

                if (collections.Entities.Count > 0)
                {
                    Entity entity = collections.Entities.First();
                    //EntityID
                    return new Guid(entity[EntityName + "id"].ToString());
                }
                else
                    return Guid.Empty;
            }
            catch (Exception ex)
            {
                EnvironmentSetting.ErrorMsg = "撈取資料失敗\n";
                EnvironmentSetting.ErrorMsg += ex.Message + "\n";
                EnvironmentSetting.ErrorMsg += ex.Source + "\n";
                EnvironmentSetting.ErrorMsg += ex.StackTrace + "\n";
                EnvironmentSetting.ErrorType = ErrorType.DATASYNCDETAIL;
                return Guid.Empty;
            }
        }
        public static Guid RetrieveEntityGuid(String linkEntityName, String linkValue, String linkFieldName, String EntityName, String value, String fieldName, String relationshipFieldName)
        {
            try
            {
                ConditionExpression condition1 = new ConditionExpression();
                condition1.AttributeName = fieldName;
                condition1.Operator = ConditionOperator.Equal;
                condition1.Values.Add(value);

                FilterExpression filter1 = new FilterExpression();
                filter1.Conditions.Add(condition1);

                ConditionExpression condition2 = new ConditionExpression();
                condition2.AttributeName = linkFieldName;
                condition2.Operator = ConditionOperator.Equal;
                condition2.Values.Add(linkValue);

                FilterExpression filter2 = new FilterExpression();
                filter2.Conditions.Add(condition2);

                LinkEntity linkEntity1 = new LinkEntity();
                linkEntity1.LinkFromEntityName = EntityName; // <--- prlot entity name
                linkEntity1.LinkFromAttributeName = relationshipFieldName;// <-- 商品代號欄位名稱
                linkEntity1.LinkToEntityName = linkEntityName;
                linkEntity1.LinkToAttributeName = linkEntityName + "id";
                linkEntity1.LinkCriteria.AddFilter(filter2);

                //EntityLogicalName
                QueryExpression qe = new QueryExpression(EntityName); // <--- prlot entity name
                qe.LinkEntities.Add(linkEntity1);
                qe.Criteria.AddFilter(filter1);

                EntityCollection collections = EnvironmentSetting.Service.RetrieveMultiple(qe);

                if (collections.Entities.Count > 0)
                {
                    Entity entity = collections.Entities.First();
                    //EntityID
                    return new Guid(entity[EntityName + "id"].ToString());
                }
                else
                    return Guid.Empty;
            }
            catch (Exception ex)
            {
                EnvironmentSetting.ErrorMsg = "撈取資料失敗\n";
                EnvironmentSetting.ErrorMsg += ex.Message + "\n";
                EnvironmentSetting.ErrorMsg += ex.Source + "\n";
                EnvironmentSetting.ErrorMsg += ex.StackTrace + "\n";
                EnvironmentSetting.ErrorType = ErrorType.DATASYNCDETAIL;
                return Guid.Empty;
            }
        }
    }
}
