using CRM.Common;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace HAT_API_CUS
{
    class CusModel
    {
        /// false no contact 
        /// true with contact
        Boolean isNeedContact = true;
        TransactionStatus transactionStatus = TransactionStatus.Success;

        //static
        private OrganizationServiceContext xrm = EnvironmentSetting.Xrm;
        private IOrganizationService service = EnvironmentSetting.Service;

        //lookup
        private int primarycontactid;
        private int new_arcontactor;
        private int new_purcontactor;
        private int new_relative_account;
        private int new_grpno;

        //int
        private int new_chkday;
        private int new_plday;

        //money
        private int creditlimit;

        //option
        private int new_collection_method;
        private int new_discount_method;
        private int new_guitype;

        //string
        private int accountnumber;
        private int address1_line1;
        private int address1_postalcode;
        private int address1_telephone2;
        private int address2_line1;
        private int fax;
        private int name;
        private int telephone1;
        private int telephone2;
        private int telephone3;
        private int new_guiid;
        private int new_guititle;
        private int new_ownid;

        //string
        private static String[] stringNameArray = { "acfaxno", "acmtelno", "inusno", "invrmk", "misno", "saleno" };
        private int[] stringIntArray = new int[stringNameArray.Length];

        public CusModel(SqlDataReader reader)
        {
            try
            {
                //special
                //lookup
                primarycontactid = reader.GetOrdinal("boss");
                new_arcontactor = reader.GetOrdinal("accman");
                new_purcontactor = reader.GetOrdinal("cman");
                new_relative_account = reader.GetOrdinal("gprna");
                new_grpno = reader.GetOrdinal("grpno");

                //int
                new_chkday = reader.GetOrdinal("chkday");
                new_plday = reader.GetOrdinal("plday");

                //money
                creditlimit = reader.GetOrdinal("credit");

                //option
                new_collection_method = reader.GetOrdinal("cpayno");
                new_discount_method = reader.GetOrdinal("cdpayno");
                new_guitype = reader.GetOrdinal("inv23");

                //string
                accountnumber = reader.GetOrdinal("asno");
                address1_line1 = reader.GetOrdinal("address");
                address1_postalcode = reader.GetOrdinal("zipno");
                address1_telephone2 = reader.GetOrdinal("mtelno");
                address2_line1 = reader.GetOrdinal("invadd");
                fax = reader.GetOrdinal("faxno");
                name = reader.GetOrdinal("asna");
                telephone1 = reader.GetOrdinal("telno");
                telephone2 = reader.GetOrdinal("bmtelno");
                telephone3 = reader.GetOrdinal("actelno");
                new_guiid = reader.GetOrdinal("invno");
                new_guititle = reader.GetOrdinal("asname");
                new_ownid = reader.GetOrdinal("idno");


                //string
                for (int i = 0, size = stringNameArray.Length; i < size; i++)
                {
                    stringIntArray[i] = reader.GetOrdinal(stringNameArray[i]);
                }

            }
            catch (Exception ex)
            {
                EnvironmentSetting.ErrorMsg += "搜尋欄位失敗\n";
                EnvironmentSetting.ErrorMsg += ex.Message + "\n";
                EnvironmentSetting.ErrorMsg += ex.Source + "\n";
                EnvironmentSetting.ErrorMsg += ex.StackTrace + "\n";
                EnvironmentSetting.ErrorType = ErrorType.DB;
            }
        }
        public Guid IsAccountExist(SqlDataReader reader)
        {
            try
            {
                return Lookup.RetrieveEntityGuid("account", reader.GetString(accountnumber).Trim(), "accountnumber");
            }
            catch (Exception ex)
            {
                EnvironmentSetting.ErrorMsg += "檢查資料失敗\n";
                EnvironmentSetting.ErrorMsg += ex.Message + "\n";
                EnvironmentSetting.ErrorMsg += ex.Source + "\n";
                EnvironmentSetting.ErrorMsg += ex.StackTrace + "\n";
                EnvironmentSetting.ErrorType = ErrorType.DATASYNCDETAIL;
                return Guid.Empty;
            }
        }
        public TransactionStatus CreateAccountForCRM(SqlDataReader reader)
        {
            return AddAttributeForRecord(reader, Guid.Empty);
        }
        public TransactionStatus UpdateAccountForCRM(SqlDataReader reader, Guid accountId)
        {
            return AddAttributeForRecord(reader, accountId);
        }
        private TransactionStatus AddAttributeForRecord(SqlDataReader reader, Guid entityId)
        {
            try
            {
                Entity entity = new Entity("account");
                //ERP撈出來寫到CRM要設定成已成交
                entity["new_businessdevelopstatus"] = new OptionSetValue(100000001);

                //int
                entity["new_chkday"] = reader.GetInt32(new_chkday);
                entity["new_plday"] = Convert.ToInt32(reader.GetInt16(new_plday));

                //string
                entity["accountnumber"] = reader.GetString(accountnumber).Trim();
                entity["address1_line1"] = reader.GetString(address1_line1).Trim();
                entity["address1_postalcode"] = reader.GetString(address1_postalcode).Trim();
                entity["address1_telephone2"] = reader.GetString(address1_telephone2).Trim();
                entity["address2_line1"] = reader.GetString(address2_line1).Trim();
                entity["fax"] = reader.GetString(fax).Trim();
                entity["name"] = reader.GetString(name).Trim();
                entity["telephone1"] = reader.GetString(telephone1).Trim();
                entity["telephone2"] = reader.GetString(telephone2).Trim();
                entity["telephone3"] = reader.GetString(telephone3).Trim();
                entity["new_guiid"] = reader.GetString(new_guiid).Trim();
                entity["new_guititle"] = reader.GetString(new_guititle).Trim();
                entity["new_ownid"] = reader.GetString(new_ownid).Trim();

                //money
                entity["creditlimit"] = new Money(reader.GetInt32(creditlimit));

                //option
                String recordStr;
                Guid recordGuid;

                recordStr = reader.GetString(new_collection_method).Trim();
                if (recordStr == "" || recordStr == null)
                    entity["new_collection_method"] = null;
                else
                    entity["new_collection_method"] = new OptionSetValue(Convert.ToInt32(recordStr));

                //recordStr = reader.GetString(new_discount_method).Trim();
                //if (recordStr == "" || recordStr == null)
                //    account["new_discount_method"] = new OptionSetValue(100000001);
                //else
                //    account["new_discount_method"] = new OptionSetValue(Convert.ToInt32(recordStr));

                //recordStr = reader.GetString(new_guitype).Trim();
                //if (recordStr == "" || recordStr == null)
                //    account["new_guitype"] = null;
                //else
                //    account["new_guitype"] = new OptionSetValue(Convert.ToInt32(recordStr) * 10);

                //string
                for (int i = 0, size = stringNameArray.Length; i < size; i++)
                {
                    entity["new_" + stringNameArray[i]] = reader.GetString(stringIntArray[i]).Trim();
                }

                //lookup
                if (isNeedContact)
                {
                    /// CRM欄位名稱     負責人     primarycontactid
                    /// CRM關聯實體     連絡人     contact
                    /// CRM關聯欄位     全名       fullname
                    /// ERP欄位名稱                boss
                    /// 
                    recordStr = reader.GetString(primarycontactid).Trim();
                    if (recordStr == "" || recordStr == null)
                        entity["primarycontactid"] = null;
                    else
                    {
                        recordGuid = Lookup.RetrieveEntityGuid("contact", recordStr, "fullname");
                        if (recordGuid == Guid.Empty)
                        {
                            EnvironmentSetting.ErrorMsg = "CRM 查無相符合資料 : \n";
                            EnvironmentSetting.ErrorMsg += "\tCRM實體 : contact\n";
                            EnvironmentSetting.ErrorMsg += "\tCRM欄位 : fullname\n";
                            EnvironmentSetting.ErrorMsg += "\tERP欄位 : boss\n";
                            //Console.WriteLine(EnvironmentSetting.ErrorMsg);
                            transactionStatus = TransactionStatus.Incomplete;
                        }
                        else
                            entity["primarycontactid"] = new EntityReference("contact", recordGuid);
                    }

                    /// CRM欄位名稱     帳務人員    new_arcontactor
                    /// CRM關聯實體     連絡人      contact
                    /// CRM關聯欄位     全名        fullname
                    /// ERP欄位名稱                 accman
                    /// 
                    recordStr = reader.GetString(new_arcontactor).Trim();
                    if (recordStr == "" || recordStr == null)
                        entity["new_arcontactor"] = null;
                    else
                    {
                        recordGuid = Lookup.RetrieveEntityGuid("contact", recordStr, "fullname");
                        if (recordGuid == Guid.Empty)
                        {
                            EnvironmentSetting.ErrorMsg = "CRM 查無相符合資料 : \n";
                            EnvironmentSetting.ErrorMsg += "\tCRM實體 : contact\n";
                            EnvironmentSetting.ErrorMsg += "\tCRM欄位 : fullname\n";
                            EnvironmentSetting.ErrorMsg += "\tERP欄位 : accman\n";
                            //EnvironmentSetting.ErrorMsg += "\tERP資料 : " + recordStr + "\n";
                            //Console.WriteLine(EnvironmentSetting.ErrorMsg);
                            transactionStatus = TransactionStatus.Incomplete;
                        }
                        entity["new_arcontactor"] = new EntityReference("contact", recordGuid);
                    }

                    /// CRM欄位名稱     訂貨聯絡人       new_purcontactor
                    /// CRM關聯實體     連絡人           contact
                    /// CRM關聯欄位     全名             fullname
                    /// ERP欄位名稱                      cman
                    /// 
                    recordStr = reader.GetString(new_purcontactor).Trim();
                    if (recordStr == "" || recordStr == null)
                        entity["new_purcontactor"] = null;
                    else
                    {
                        recordGuid = Lookup.RetrieveEntityGuid("contact", recordStr, "fullname");
                        if (recordGuid == Guid.Empty)
                        {
                            EnvironmentSetting.ErrorMsg = "CRM 查無相符合資料 : \n";
                            EnvironmentSetting.ErrorMsg += "\tCRM實體 : contact\n";
                            EnvironmentSetting.ErrorMsg += "\tCRM欄位 : fullname\n";
                            EnvironmentSetting.ErrorMsg += "\tERP欄位 : cman\n";
                            //EnvironmentSetting.ErrorMsg += "\tERP資料 : " + recordStr + "\n";
                            //Console.WriteLine(EnvironmentSetting.ErrorMsg);
                            transactionStatus = TransactionStatus.Incomplete;
                        }
                        entity["new_purcontactor"] = new EntityReference("contact", recordGuid);
                    }
                }
                /// CRM欄位名稱     外釋診所    new_relative_account
                /// CRM關聯實體     客戶        account
                /// CRM關聯欄位     客戶名稱    name
                /// ERP欄位名稱                 gprna
                /// 
                recordStr = reader.GetString(new_relative_account).Trim();
                if (recordStr == "" || recordStr == null)
                    entity["new_relative_account"] = null;
                else
                {
                    recordGuid = Lookup.RetrieveEntityGuid("account", recordStr, "name");
                    if (recordGuid == Guid.Empty)
                    {
                        EnvironmentSetting.ErrorMsg = "CRM 查無相符合資料 : \n";
                        EnvironmentSetting.ErrorMsg += "\tCRM實體 : account\n";
                        EnvironmentSetting.ErrorMsg += "\tCRM欄位 : name\n";
                        EnvironmentSetting.ErrorMsg += "\tERP欄位 : gprna\n";
                        //EnvironmentSetting.ErrorMsg += "\tERP資料 : " + recordStr + "\n";
                        //Console.WriteLine(EnvironmentSetting.ErrorMsg);
                        transactionStatus= TransactionStatus.Incomplete;
                    }
                    entity["new_relative_account"] = new EntityReference("account", recordGuid);
                }

                /// CRM欄位名稱     醫院群組        new_grpno
                /// CRM關聯實體     客戶醫院群組    new_grpno
                /// CRM關聯欄位     醫院群組        new_grpno
                /// ERP欄位名稱                     gprna
                /// 
                recordStr = reader.GetString(new_grpno).Trim();
                if (recordStr == "" || recordStr == null)
                    entity["new_grpno"] = null;
                else
                {
                    recordGuid = Lookup.RetrieveEntityGuid("new_grpno", recordStr, "new_grpno");
                    if (recordGuid == Guid.Empty)
                    {
                        EnvironmentSetting.ErrorMsg = "CRM 查無相符合資料 : \n";
                        EnvironmentSetting.ErrorMsg += "\tCRM實體 : new_grpno\n";
                        EnvironmentSetting.ErrorMsg += "\tCRM欄位 : new_grpno\n";
                        EnvironmentSetting.ErrorMsg += "\tERP欄位 : grpno\n";
                        //EnvironmentSetting.ErrorMsg += "\tERP資料 : " + recordStr + "\n";
                        //Console.WriteLine(EnvironmentSetting.ErrorMsg);
                        transactionStatus =  TransactionStatus.Incomplete;
                    }
                    entity["new_grpno"] = new EntityReference("new_grpno", recordGuid);
                }

                //[新增] OR [更新] 資料
                try
                {
                    if (entityId == Guid.Empty)
                        service.Create(entity);
                    else
                    {
                        entity["accountid"] = entityId;
                        service.Update(entity);
                    }
                    return transactionStatus;
                }
                catch (Exception ex)
                {
                    //Console.WriteLine("accountnumber : " + reader.GetString(accountnumber).Trim());
                    //Console.WriteLine(ex.Message);
                    EnvironmentSetting.ErrorMsg = ex.Message;
                    return TransactionStatus.Fail;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("欄位讀取錯誤");
                Console.WriteLine(ex.Message);
                EnvironmentSetting.ErrorMsg = "欄位讀取錯誤\n" + ex.Message;
                return TransactionStatus.Fail;
            }
        }
    }
}
