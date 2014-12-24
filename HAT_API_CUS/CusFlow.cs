using CRM.Common;
using CRM.Model;
using NLog;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace HAT_API_CUS
{
    class CusFlow
    {
        private static Logger _logger = LogManager.GetCurrentClassLogger();
        public void DoCusFlow()
        {
            DataSyncModel dataSync = new DataSyncModel();
            //connection CRM
            EnvironmentSetting.LoadSetting();
            if (EnvironmentSetting.ErrorType == ErrorType.None)
            {
                //create dataSync
                dataSync.CreateDataSyncForCRM("客戶");
                if (EnvironmentSetting.ErrorType == ErrorType.None)
                {
                    //connection DB
                    EnvironmentSetting.LoadDB();
                    if (EnvironmentSetting.ErrorType == ErrorType.None)
                    {
                        //get information from DB
                        EnvironmentSetting.GetList("select * from dbo.v_cus");
                        if (EnvironmentSetting.ErrorType == ErrorType.None)
                        {
                            //get reader
                            SqlDataReader reader = EnvironmentSetting.Reader;

                            //search field index
                            CusModel cusmodel = new CusModel(reader);
                            if (EnvironmentSetting.ErrorType == ErrorType.None)
                            {
                                Console.WriteLine("連線成功!!");
                                Console.WriteLine("開始執行...");

                                if (reader.HasRows)
                                {
                                    int success = 0;
                                    int fail = 0;
                                    while (reader.Read())
                                    {
                                        //判斷CRM是否有資料
                                        Guid existAccountId = cusmodel.IsAccountExist(EnvironmentSetting.Reader);
                                        if (EnvironmentSetting.ErrorType == ErrorType.None)
                                        {
                                            TransactionStatus transactionStatus;
                                            TransactionType transactionType;
                                            if (existAccountId == Guid.Empty)
                                            {
                                                //create
                                                transactionType = TransactionType.Insert;
                                                transactionStatus = cusmodel.CreateAccountForCRM(EnvironmentSetting.Reader);
                                            }
                                            else
                                            {
                                                //update
                                                transactionType = TransactionType.Update;
                                                transactionStatus = cusmodel.UpdateAccountForCRM(EnvironmentSetting.Reader, existAccountId);
                                            }

                                            if (EnvironmentSetting.ErrorType == ErrorType.None)
                                            {
                                                //create datasyncdetail
                                                switch (transactionStatus)
                                                {
                                                    case TransactionStatus.Success:
                                                        success++;
                                                        break;
                                                    case TransactionStatus.Fail:
                                                        //新增、更新資料有錯誤 則新增一筆detail
                                                        dataSync.CreateDataSyncDetailForCRM(reader["asno"].ToString().Trim(), reader["asna"].ToString().Trim(), transactionType, transactionStatus);
                                                        fail++;
                                                        break;
                                                    default:
                                                        fail++;
                                                        break;
                                                }

                                                //新增detail錯誤 則結束
                                                if (EnvironmentSetting.ErrorType != ErrorType.None)
                                                {
                                                    //_logger.Info("asno" + reader["asno"].ToString().Trim());
                                                    //_logger.Info("asna" + reader["asna"].ToString().Trim());
                                                    //Console.WriteLine("asno" + reader["asno"].ToString().Trim());
                                                    //Console.WriteLine("asna" + reader["asna"].ToString().Trim());
                                                    //Console.WriteLine();
                                                    _logger.Info(EnvironmentSetting.ErrorMsg);
                                                    EnvironmentSetting.ErrorType = ErrorType.None;
                                                    //break;
                                                }
                                            }
                                        }
                                    }
                                    //更新DataSync 成功、失敗、完成時間
                                    dataSync.UpdateDataSyncForCRM(success, fail);
                                }
                                else
                                {
                                    Console.WriteLine("沒有資料");
                                    EnvironmentSetting.ErrorMsg += "ERP沒有資料\n";
                                    EnvironmentSetting.ErrorType = ErrorType.DB;
                                }
                            }
                        }
                    }
                }
            }
            switch (EnvironmentSetting.ErrorType)
            {
                case ErrorType.None:
                    break;

                case ErrorType.INI:
                case ErrorType.CRM:
                case ErrorType.DATASYNC:
                    _logger.Info(EnvironmentSetting.ErrorMsg);
                    break;

                case ErrorType.DB:
                case ErrorType.DATASYNCDETAIL:
                    dataSync.UpdateDataSyncWithErrorForCRM(EnvironmentSetting.ErrorMsg);
                    break;

                default:
                    break;
            }
            Console.WriteLine("執行完畢...");
            Console.ReadLine();
        }
    }
}
