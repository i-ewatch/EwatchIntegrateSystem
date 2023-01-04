using Dapper;
using EwatchIntegrateSystem.Enums;
using EwatchIntegrateSystem.Modules;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Data.SqlClient;

namespace EwatchIntegrateSystem.Methods
{
    /// <summary>
    /// 電表資訊資料庫方法
    /// </summary>
    public class ElectricLogMethod
    {
        private string SqlDB { get; set; }
        /// <summary>
        /// 錯誤資訊
        /// </summary>
        public string ErrorResponse { get; set; } = "";
        /// <summary>
        /// 電表資訊資料庫方法初始化
        /// </summary>
        /// <param name="sqlDB"></param>
        public ElectricLogMethod(string sqlDB)
        {
            SqlDB = sqlDB;
        }
        #region 建立資料庫與資料表
        /// <summary>
        /// 建立資料庫
        /// </summary>
        /// <param name="dateTime"></param>
        private void Create_DataBase(DateTime dateTime)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    //string sql = $"SELECT COUNT(*) FROM  master.dbo.sysdatabases WHERE name='EwatchIntegrate_Log_{dateTime:yyyyMM}'";
                    string sql = $"SELECT COUNT(*) FROM  master.dbo.sysdatabases WHERE name='EwatchIntegrate_Log'";
                    int databaseflag = connection.QuerySingle<int>(sql);
                    if (databaseflag == 0)
                    {
                        //string databaseSql = $"CREATE DATABASE [EwatchIntegrate_Log_{dateTime:yyyyMM}]";
                        string databaseSql = $"CREATE DATABASE [EwatchIntegrate_Log]";
                        connection.Execute(databaseSql);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
            }
        }
        /// <summary>
        /// 建立電表資料表
        /// </summary>
        /// <param name="dateTime"></param>
        /// <param name="CardNumber"></param>
        /// <param name="BordNumber"></param>
        private void Create_ElectricTable(DateTime dateTime, string CardNumber, string BordNumber)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    //string sql = $"USE [EwatchIntegrate_Log_{dateTime:yyyyMM}] SELECT COUNT(*) FROM  sys.tables WHERE name='Electric_{CardNumber}{BordNumber}'";
                    string sql = $"USE [EwatchIntegrate_Log] SELECT COUNT(*) FROM  sys.tables WHERE name='Electric_{CardNumber}{BordNumber}'";
                    int Tableflag = connection.QuerySingle<int>(sql);
                    if (Tableflag == 0)
                    {
                        //string dataTableSql = $"USE [EwatchIntegrate_Log_{dateTime:yyyyMM}] CREATE TABLE [Electric_{CardNumber}{BordNumber}] (" +
                        string dataTableSql = $"USE [EwatchIntegrate_Log] CREATE TABLE [Electric_{CardNumber}{BordNumber}] (" +
                            $" [ttime] NVARCHAR(14) NOT NULL" +
                            $",[ttimen] DATETIME NOT NULL" +
                            $",[DeviceNumber] Int NOT NULL" +
                            $",[Loop] INT NOT NULL DEFAULT 0" +
                            $",[Va] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Vb] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Vc] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Vnavg] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Vab] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Vbc] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Vca] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Vlavg] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Ia] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Ib] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Ic] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Iavg] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Freq] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PFa] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PFb] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PFc] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PFavg] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Kwa] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Kwb] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Kwc] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Kwtot] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[KVARa] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[KVARb] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[KVARc] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[KVARtot] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[KVAa] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[KVAb] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[KVAc] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[KVAtot] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PhaseAngleVa] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PhaseAngleVb] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PhaseAngleVc] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PhaseAngleIa] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PhaseAngleIb] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PhaseAngleIc] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD1] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD2] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD3] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD4] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD5] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD6] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD7] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD8] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD9] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD10] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD11] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD12] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD13] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD14] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD15] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD16] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD17] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD18] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD19] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD20] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD21] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD22] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD23] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD24] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD25] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD26] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD27] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD28] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD29] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD30] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RV_HD31] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD1] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD2] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD3] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD4] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD5] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD6] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD7] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD8] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD9] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD10] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD11] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD12] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD13] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD14] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD15] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD16] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD17] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD18] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD19] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD20] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD21] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD22] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD23] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD24] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD25] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD26] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD27] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD28] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD29] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD30] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SV_HD31] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD1] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD2] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD3] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD4] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD5] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD6] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD7] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD8] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD9] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD10] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD11] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD12] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD13] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD14] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD15] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD16] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD17] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD18] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD19] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD20] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD21] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD22] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD23] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD24] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD25] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD26] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD27] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD28] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD29] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD30] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TV_HD31] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD1] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD2] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD3] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD4] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD5] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD6] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD7] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD8] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD9] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD10] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD11] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD12] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD13] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD14] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD15] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD16] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD17] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD18] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD19] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD20] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD21] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD22] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD23] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD24] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD25] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD26] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD27] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD28] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD29] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD30] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[RI_HD31] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD1] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD2] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD3] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD4] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD5] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD6] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD7] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD8] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD9] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD10] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD11] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD12] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD13] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD14] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD15] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD16] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD17] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD18] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD19] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD20] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD21] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD22] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD23] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD24] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD25] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD26] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD27] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD28] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD29] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD30] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[SI_HD31] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD1] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD2] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD3] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD4] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD5] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD6] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD7] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD8] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD9] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD10] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD11] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD12] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD13] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD14] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD15] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD16] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD17] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD18] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD19] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD20] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD21] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD22] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD23] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD24] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD25] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD26] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD27] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD28] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD29] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD30] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[TI_HD31] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[KWH] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",CONSTRAINT [PK_Electric_{CardNumber}{BordNumber}] PRIMARY KEY ([ttime], [ttimen], [DeviceNumber], [Loop]))";
                        connection.Execute(dataTableSql);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
            }
        }
        /// <summary>
        /// 建立電表驟降事件表(不使用)
        /// </summary>
        /// <param name="dateTime"></param>
        /// <param name="CardNumber"></param>
        /// <param name="BordNumber"></param>
        private void Create_ElectricSagTable(DateTime dateTime, string CardNumber, string BordNumber)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = $"USE [EwatchIntegrate_Log_{dateTime:yyyyMM}] SELECT COUNT(*) FROM  sys.tables WHERE name='ElectricSag_{CardNumber}{BordNumber}'";
                    int Tableflag = connection.QuerySingle<int>(sql);
                    if (Tableflag == 0)
                    {
                        string dataTableSql = $"USE [EwatchIntegrate_Log_{dateTime:yyyyMM}] CREATE TABLE [ElectricSag_{CardNumber}{BordNumber}] (" +
                            $" [ttime] NVARCHAR(16) NOT NULL" +
                            $",[ttimen] DATETIME NOT NULL" +
                            $",[DeviceNumber] Int NOT NULL" +
                            $",[sagCycle] INT NOT NULL DEFAULT 0" +
                            $",[sagData] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[sagPhase] NVARCHAR(2) NOT NULL DEFAULT ''" +
                            $",[sagBegin] NVARCHAR(17) NOT NULL DEFAULT ''" +
                            $",[sagEnd] NVARCHAR(17) NOT NULL DEFAULT ''" +
                            $",CONSTRAINT [PK_ElectricSag_{CardNumber}{BordNumber}] PRIMARY KEY ([ttime], [ttimen], [DeviceNumber]))";
                        connection.Execute(dataTableSql);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
            }
        }
        /// <summary>
        /// 建立電表驟升事件表(不使用)
        /// </summary>
        /// <param name="dateTime"></param>
        /// <param name="CardNumber"></param>
        /// <param name="BordNumber"></param>
        private void Create_ElectricSwellTable(DateTime dateTime, string CardNumber, string BordNumber)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = $"USE [EwatchIntegrate_Log_{dateTime:yyyyMM}] SELECT COUNT(*) FROM  sys.tables WHERE name='ElectricSwell_{CardNumber}{BordNumber}'";
                    int Tableflag = connection.QuerySingle<int>(sql);
                    if (Tableflag == 0)
                    {
                        string dataTableSql = $"USE [EwatchIntegrate_Log_{dateTime:yyyyMM}] CREATE TABLE [ElectricSwell_{CardNumber}{BordNumber}] (" +
                            $" [ttime] NVARCHAR(16) NOT NULL" +
                            $",[ttimen] DATETIME NOT NULL" +
                            $",[DeviceNumber] Int NOT NULL" +
                            $",[swellCycle] INT NOT NULL DEFAULT 0" +
                            $",[swellData] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[swellPhase] NVARCHAR(2) NOT NULL DEFAULT ''" +
                            $",[swellBegin] NVARCHAR(17) NOT NULL DEFAULT ''" +
                            $",[swellEnd] NVARCHAR(17) NOT NULL DEFAULT ''" +
                            $",CONSTRAINT [PK_ElectricSwell_{CardNumber}{BordNumber}] PRIMARY KEY ([ttime], [ttimen], [DeviceNumber]))";
                        connection.Execute(dataTableSql);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
            }
        }
        /// <summary>
        /// 建立電表驟降驟升事件表
        /// </summary>
        /// <param name="dateTime"></param>
        /// <param name="CardNumber"></param>
        /// <param name="BordNumber"></param>
        private void Create_Electric_Sag_SwellTable(DateTime dateTime, string CardNumber, string BordNumber)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    //string sql = $"USE [EwatchIntegrate_Log_{dateTime:yyyyMM}] SELECT COUNT(*) FROM  sys.tables WHERE name='ElectricSwell_{CardNumber}{BordNumber}'";
                    string sql = $"USE [EwatchIntegrate_Log] SELECT COUNT(*) FROM  sys.tables WHERE name='Electric_Sag_Swell_{CardNumber}{BordNumber}'";
                    int Tableflag = connection.QuerySingle<int>(sql);
                    if (Tableflag == 0)
                    {
                        //string dataTableSql = $"USE [EwatchIntegrate_Log_{dateTime:yyyyMM}] CREATE TABLE [ElectricSwell_{CardNumber}{BordNumber}] (" +
                        string dataTableSql = $"USE [EwatchIntegrate_Log] CREATE TABLE [Electric_Sag_Swell_{CardNumber}{BordNumber}] (" +
                            $" [ttime] NVARCHAR(16) NOT NULL" +
                            $",[ttimen] DATETIME NOT NULL" +
                            $",[DeviceNumber] Int NOT NULL" +
                            $",[Sag_Swell_Type] NVARCHAR(2) NOT NULL DEFAULT ''" +
                            $",[Cycle] INT NOT NULL DEFAULT 0" +
                            $",[Data] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[Phase] NVARCHAR(2) NOT NULL DEFAULT ''" +
                            $",[Sag_Swell_Begin] NVARCHAR(17) NOT NULL DEFAULT ''" +
                            $",[Sag_Swell_End] NVARCHAR(17) NOT NULL DEFAULT ''" +
                            $",CONSTRAINT [PK_ElectricSwell_{CardNumber}{BordNumber}] PRIMARY KEY ([ttime], [ttimen], [DeviceNumber]))";
                        connection.Execute(dataTableSql);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
            }
        }
        /// <summary>
        /// 建立電表告警事件表
        /// </summary>
        /// <param name="dateTime"></param>
        /// <param name="CardNumber"></param>
        /// <param name="BordNumber"></param>
        private void Create_ElectricAlarmTable(DateTime dateTime, string CardNumber, string BordNumber)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    //string sql = $"USE [EwatchIntegrate_Log_{dateTime:yyyyMM}] SELECT COUNT(*) FROM  sys.tables WHERE name='ElectricAlarm_{CardNumber}{BordNumber}'";
                    string sql = $"USE [EwatchIntegrate_Log] SELECT COUNT(*) FROM  sys.tables WHERE name='ElectricAlarm_{CardNumber}{BordNumber}'";
                    int Tableflag = connection.QuerySingle<int>(sql);
                    if (Tableflag == 0)
                    {
                        string dataTableSql = $"USE [EwatchIntegrate_Log] CREATE TABLE [ElectricAlarm_{CardNumber}{BordNumber}] (" +
                            $" [ttime] NVARCHAR(16) NOT NULL" +
                            $",[ttimen] DATETIME NOT NULL" +
                            $",[DeviceNumber] Int NOT NULL" +
                            $",[AlarmItem] NVARCHAR(50) NOT NULL DEFAULT ''" +
                            $",[AlarmData] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[AlarmTime] NVARCHAR(17) NOT NULL DEFAULT ''" +
                            $",CONSTRAINT [PK_ElectricAlarm_{CardNumber}{BordNumber}] PRIMARY KEY ([ttime], [ttimen], [DeviceNumber]))";
                        connection.Execute(dataTableSql);
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
            }
        }
        #endregion
        #region 卡版號是否存在
        /// <summary>
        /// 卡版號是否存在
        /// </summary>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        public bool CardBord_Search(string CardNumber, string BordNumber)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = $"SELECT COUNT(*) FROM CaseSetting  WHERE CardNumber = @CardNumber AND BordNumber = @BordNumber";
                    int Index = connection.QuerySingle<int>(sql, new { CardNumber = CardNumber, BordNumber = BordNumber });
                    if (Index > 0)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return false;
            }
        }
        #endregion
        #region 單顆電表查詢資訊
        /// <summary>
        /// 月用電趨勢
        /// </summary>
        /// <param name="StartTime">查詢時間 yyyy/MM</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <param name="DeviceName">設備名稱</param>
        /// <param name="DeviceNumber">設備編號</param>
        /// <returns></returns>
        public DataTable Search_Month_kWh(DateTime StartTime, string CardNumber, string BordNumber, string DeviceName, int DeviceNumber)
        {
            try
            {
                DateTime nowstartTime = Convert.ToDateTime(StartTime.ToString("yyyy/MM/01 00:00:00"));
                DateTime nowendTime = Convert.ToDateTime(nowstartTime.ToString("yyyy/MM") + $"/{DateTime.DaysInMonth(nowstartTime.Year, nowstartTime.Month).ToString().PadLeft(2, '0')}");
                DateTime afterstartTime = nowstartTime.AddYears(-1);
                DateTime afterendTime = Convert.ToDateTime(afterstartTime.ToString("yyyy/MM") + $"/{DateTime.DaysInMonth(afterstartTime.Year, afterstartTime.Month).ToString().PadLeft(2, '0')}");
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string TempTableData = string.Empty;
                    string ColumnsData = string.Empty;
                    string LeftJoinCommand = string.Empty;
                    DataTable dt = new DataTable();
                    dt.Columns.Add("ttime", typeof(string));
                    dt.Columns.Add($"NowTotalkWh", typeof(string));
                    dt.Columns.Add($"AfterTotalkWh", typeof(string));
                    TempTableData += GetTempTableSingleData(nowstartTime, nowendTime, CardNumber, BordNumber, DeviceNumber, false, false, "Total");
                    TempTableData += GetTempTableSingleData(afterstartTime, afterendTime, CardNumber, BordNumber, DeviceNumber, true, false, "Total");
                    ColumnsData += GetSingleColumnsData(DeviceNumber, DeviceName, "NowTotalkWh", false);
                    ColumnsData += GetSingleColumnsData(DeviceNumber, DeviceName, "AfterTotalkWh", true);
                    LeftJoinCommand += GetLeftJoinCommand(DeviceNumber, false);
                    LeftJoinCommand += GetLeftJoinCommand(DeviceNumber, true);
                    string sqlLog = $"USE [EwatchIntegrateTotal] DECLARE @regTIme DATETIME, @startTime DATETIME, @endTime DATETIME " +
                                 $"SET @startTime ='{nowstartTime:yyyy/MM/dd}' " +
                                 $"SET @endTime = '{nowendTime:yyyy/MM/dd}' " +
                                 $"SET @regTIme = @startTime " +
                                 $"DECLARE @mainTemp AS TABLE (ttime CHAR(5),ttimen DATETIME) " +
                                 $"WHILE @regTIme <= @endTime " +
                                 $"BEGIN " +
                                 $"INSERT INTO @mainTemp(ttime, ttimen) VALUES (FORMAT(@regTIme, N'MM/dd'), @regTIme) " +
                                 $"SET @regTIme = DATEADD(DAY,1,@regTIme) " +
                                 $"END " +
                                 $"{TempTableData} " +
                                 $"SELECT  mainT.ttime AS [ttime]{ColumnsData} FROM @mainTemp AS mainT {LeftJoinCommand}";
                    var dr = connection.ExecuteReader(sqlLog);
                    dt.Load(dr);
                    return dt;
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return new();
            }
        }
        /// <summary>
        /// 年用電趨勢
        /// </summary>
        /// <param name="StartTime">查詢時間 yyyy/MM</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <param name="DeviceName">設備名稱</param>
        /// <param name="DeviceNumber">設備編號</param>
        /// <returns></returns>
        public DataTable Search_Year_kWh(DateTime StartTime, string CardNumber, string BordNumber, string DeviceName, int DeviceNumber)
        {
            try
            {
                DateTime nowstartTime = Convert.ToDateTime(StartTime.ToString("yyyy/01/01 00:00:00"));
                DateTime nowendTime = Convert.ToDateTime(nowstartTime.ToString("yyyy/12/31  00:00:00"));
                DateTime afterstartTime = nowstartTime.AddYears(-1);
                DateTime afterendTime = Convert.ToDateTime(afterstartTime.ToString("yyyy/12/31 00:00:00"));
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string TempTableData = string.Empty;
                    string ColumnsData = string.Empty;
                    string LeftJoinCommand = string.Empty;
                    DataTable dt = new DataTable();
                    dt.Columns.Add("ttime", typeof(string));
                    dt.Columns.Add($"NowTotalkWh", typeof(string));
                    dt.Columns.Add($"AfterTotalkWh", typeof(string));
                    TempTableData += GetTempTableSingleData(nowstartTime, nowendTime, CardNumber, BordNumber, DeviceNumber, false, true, "Total");
                    TempTableData += GetTempTableSingleData(afterstartTime, afterendTime, CardNumber, BordNumber, DeviceNumber, true, true, "Total");
                    ColumnsData += GetSingleColumnsData(DeviceNumber, DeviceName, "NowTotalkWh", false);
                    ColumnsData += GetSingleColumnsData(DeviceNumber, DeviceName, "AfterTotalkWh", true);
                    LeftJoinCommand += GetLeftJoinCommand(DeviceNumber, false);
                    LeftJoinCommand += GetLeftJoinCommand(DeviceNumber, true);
                    string sqlLog = $"USE [EwatchIntegrateTotal] DECLARE @regTIme DATETIME, @startTime DATETIME, @endTime DATETIME " +
                                 $"SET @startTime ='{nowstartTime:yyyy/MM/dd}' " +
                                 $"SET @endTime = '{nowendTime:yyyy/MM/dd}' " +
                                 $"SET @regTIme = @startTime " +
                                 $"DECLARE @mainTemp AS TABLE (ttime CHAR(2),ttimen DATETIME) " +
                                 $"WHILE @regTIme <= @endTime " +
                                 $"BEGIN " +
                                 $"INSERT INTO @mainTemp(ttime, ttimen) VALUES (FORMAT(@regTIme, N'MM'), @regTIme) " +
                                 $"SET @regTIme = DATEADD(MONTH,1,@regTIme) " +
                                 $"END " +
                                 $"{TempTableData} " +
                                 $"SELECT  mainT.ttime AS [ttime]{ColumnsData} FROM @mainTemp AS mainT {LeftJoinCommand}";
                    var dr = connection.ExecuteReader(sqlLog);
                    dt.Load(dr);
                    return dt;
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return new();
            }
        }
        private string GetTempTableSingleData(DateTime start, DateTime end, string CardNumber, string BordNumber, int DeviceNumber, bool flag, bool Year_MonthFalg, string val)
        {
            if (Year_MonthFalg)
            {
                if (!flag)//本年
                {
                    return $"DECLARE @temp{DeviceNumber} AS TABLE (" +
                           $"ttime CHAR(2), ttimen CHAR(2), val DECIMAL(18,2)) INSERT INTO @temp{DeviceNumber} (ttime,ttimen,val) " +
                           $"SELECT MAX(SUBSTRING(ttime,5,2)) as ttime,MAX(FORMAT( ttimen,N'MM')) AS ttimen, SUM({val}) AS val FROM Electric_Log WHERE CardNumber='{CardNumber}' AND BordNumber='{BordNumber}' AND DeviceNumber = '{DeviceNumber}' AND ttime >= '{start:yyyyMMdd}' AND ttime <= '{end:yyyyMMdd}' GROUP BY SUBSTRING(ttime,5,2) ORDER BY ttime ";
                }
                else//去年
                {
                    return $"DECLARE @temp{DeviceNumber + 1} AS TABLE (" +
                           $"ttime CHAR(2), ttimen CHAR(2), val DECIMAL(18,2)) INSERT INTO @temp{DeviceNumber + 1} (ttime,ttimen,val) " +
                           $"SELECT MAX(SUBSTRING(ttime,5,2)) as ttime,MAX(FORMAT( ttimen,N'MM')) AS ttimen, SUM({val}) AS val FROM Electric_Log WHERE CardNumber='{CardNumber}' AND BordNumber='{BordNumber}' AND DeviceNumber = '{DeviceNumber}' AND ttime >= '{start:yyyyMMdd}' AND ttime <= '{end:yyyyMMdd}' GROUP BY SUBSTRING(ttime,5,2) ORDER BY ttime ";
                }
            }
            else
            {
                if (!flag)//本年本月
                {
                    return $"DECLARE @temp{DeviceNumber} AS TABLE (" +
                           $"ttime CHAR(5), ttimen CHAR(5), val DECIMAL(18,2)) INSERT INTO @temp{DeviceNumber} (ttime,ttimen,val) " +
                           $"SELECT SUBSTRING(ttime,5,4) as ttime,FORMAT( ttimen,N'MM/dd') AS ttimen, {val} AS val FROM Electric_Log WHERE CardNumber='{CardNumber}' AND BordNumber='{BordNumber}' AND DeviceNumber = '{DeviceNumber}' AND ttime >= '{start:yyyyMMdd}' AND ttime <= '{end:yyyyMMdd}' ORDER BY ttime ";
                }
                else//去年本月
                {
                    return $"DECLARE @temp{DeviceNumber + 1} AS TABLE (" +
                           $"ttime CHAR(5), ttimen CHAR(5), val DECIMAL(18,2)) INSERT INTO @temp{DeviceNumber + 1} (ttime,ttimen,val) " +
                           $"SELECT SUBSTRING(ttime,5,4) as ttime,FORMAT( ttimen,N'MM/dd') AS ttimen, {val} AS val FROM Electric_Log WHERE CardNumber='{CardNumber}' AND BordNumber='{BordNumber}' AND DeviceNumber = '{DeviceNumber}' AND ttime >= '{start:yyyyMMdd}' AND ttime <= '{end:yyyyMMdd}' ORDER BY ttime ";
                }
            }
        }
        private string GetSingleColumnsData(int DeviceNumber, string DeviceName, string valName, bool flag)
        {
            if (!flag)//本年本月
            {
                return $",T{DeviceNumber}.val AS [{valName}]";
            }
            else//去年本月
            {
                return $",T{DeviceNumber + 1}.val AS [{valName}]";
            }
        }
        private string GetLeftJoinCommand(int DeviceNumber, bool flag)
        {
            if (!flag)//本年本月
            {
                return $" LEFT JOIN @temp{DeviceNumber} AS T{DeviceNumber} ON mainT.ttime = T{DeviceNumber}.ttimen";
            }
            else//去年本月
            {
                return $" LEFT JOIN @temp{DeviceNumber + 1} AS T{DeviceNumber + 1} ON mainT.ttime = T{DeviceNumber + 1}.ttimen";
            }
        }
        /// <summary>
        /// 用電累積量資訊
        /// </summary>
        /// <param name="StartTime">查詢時間 yyyy/MM</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <param name="DeviceName">設備名稱</param>
        /// <param name="DeviceNumber">設備編號</param>
        /// <returns></returns>
        public object Search_Total_Info(DateTime StartTime, string CardNumber, string BordNumber, string DeviceName, int DeviceNumber)
        {
            object T = new object();
            try
            {
                DateTime nowdaystartTime = Convert.ToDateTime(StartTime.ToString("yyyy/MM/dd 00:00:00"));
                DateTime afterdaystartTime = nowdaystartTime.AddDays(-1);
                DateTime afterMonthstartTime = nowdaystartTime.AddMonths(-1);
                DateTime afterYearstartTime = nowdaystartTime.AddYears(-1);
                List<ElectricTotal_Log> NowDayTotal = new List<ElectricTotal_Log>();
                List<ElectricTotal_Log> AfterDayTotal = new List<ElectricTotal_Log>();
                List<ElectricTotal_Log> NowMonthTotal = new List<ElectricTotal_Log>();
                List<ElectricTotal_Log> AfterMonthTotal = new List<ElectricTotal_Log>();
                List<Co2Setting> co2Settings = new List<Co2Setting>();
                decimal NowDayTotalvalue = 0;
                decimal AfterDayTotalvalue = 0;
                decimal NowMonthTotalvalue = 0;
                decimal AfterMonthTotalvalue = 0;
                decimal NowYearTotalvalue = 0;
                decimal AfterYearTotalvalue = 0;
                decimal Percent = 0;
                ElectricWeb ElectricWeb = new ElectricWeb();
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = $"USE [EwatchIntegrateTotal] SELECT * FROM Electric_Log WHERE CardNumber='{CardNumber}' AND BordNumber='{BordNumber}' AND DeviceNumber = '{DeviceNumber}' AND ttime = @start";
                    NowDayTotal = connection.Query<ElectricTotal_Log>(sql, new { start = nowdaystartTime.ToString("yyyyMMdd") }).ToList();
                    Thread.Sleep(10);
                    AfterDayTotal = connection.Query<ElectricTotal_Log>(sql, new { start = afterdaystartTime.ToString("yyyyMMdd") }).ToList();
                    sql = $"USE [EwatchIntegrateTotal] SELECT  MAX(SUBSTRING(ttime,5,2)) as ttime,MAX(FORMAT( ttimen,N'yyyy/MM')) AS ttimen, SUM(Total) AS Total FROM Electric_Log WHERE CardNumber='{CardNumber}' AND BordNumber='{BordNumber}' AND DeviceNumber = '{DeviceNumber}' AND ttime >= @start AND ttime <= @end GROUP BY SUBSTRING(ttime,5,2)";
                    NowMonthTotal = connection.Query<ElectricTotal_Log>(sql, new { start = nowdaystartTime.ToString("yyyyMM01"), end = nowdaystartTime.ToString("yyyyMM") + $"{DateTime.DaysInMonth(nowdaystartTime.Year, nowdaystartTime.Month).ToString().PadLeft(2, '0')}" }).ToList();
                    Thread.Sleep(10);
                    AfterMonthTotal = connection.Query<ElectricTotal_Log>(sql, new { start = afterMonthstartTime.ToString("yyyyMM01"), end = afterMonthstartTime.ToString("yyyyMM") + $"{DateTime.DaysInMonth(afterMonthstartTime.Year, afterMonthstartTime.Month).ToString().PadLeft(2, '0')}" }).ToList();
                    sql = $"SELECT  Kwtot * 100.0 / (CASE( SELECT SUM ( Kwtot ) FROM ElectricWeb WHERE CardNumber='{CardNumber}' AND BordNumber='{BordNumber}' )WHEN 0 THEN 100 ELSE (SELECT SUM ( Kwtot ) FROM ElectricWeb WHERE CardNumber='{CardNumber}' AND BordNumber='{BordNumber}') END) as percent_Info From ElectricWeb WHERE CardNumber='{CardNumber}' AND BordNumber='{BordNumber}' AND DeviceNumber = '{DeviceNumber}'";
                    Percent = connection.QuerySingle<decimal>(sql);
                    sql = $"USE [EwatchIntegrateTotal] SELECT SUM(Total) AS Total FROM Electric_Log WHERE CardNumber='{CardNumber}' AND BordNumber='{BordNumber}' AND DeviceNumber = '{DeviceNumber}' AND ttime >= @start AND ttime <= @end GROUP BY SUBSTRING(ttime,1,4)";
                    NowYearTotalvalue = connection.QuerySingleOrDefault<decimal>(sql, new { start = nowdaystartTime.ToString("yyyy0101"), end = nowdaystartTime.ToString("yyyy1231") });
                    Thread.Sleep(10);
                    AfterYearTotalvalue = connection.QuerySingleOrDefault<decimal>(sql, new { start = afterYearstartTime.ToString("yyyy0101"), end = afterYearstartTime.ToString("yyyy1231") });
                    sql = $"SELECT * FROM Co2Setting";
                    co2Settings = connection.Query<Co2Setting>(sql).ToList();
                    sql = $"SELECT Vab,Vbc,Vca,Ia,Ib,Ic,Kwtot,KVARtot,KVAtot,PFavg FROM ElectricWeb WHERE CardNumber='{CardNumber}' AND BordNumber='{BordNumber}' AND DeviceNumber = '{DeviceNumber}'";
                    ElectricWeb = connection.QuerySingleOrDefault<ElectricWeb>(sql);
                    if (NowDayTotal.Count > 0)
                    {
                        NowDayTotalvalue = NowDayTotal[0].Total;
                    }
                    if (AfterDayTotal.Count > 0)
                    {
                        AfterDayTotalvalue = AfterDayTotal[0].Total;
                    }
                    if (NowMonthTotal.Count > 0)
                    {
                        NowMonthTotalvalue = NowMonthTotal[0].Total;
                    }
                    if (AfterMonthTotal.Count > 0)
                    {
                        AfterMonthTotalvalue = AfterMonthTotal[0].Total;
                    }
                    decimal YearCO2Total = 0;//今年碳排
                    var nowco2 = co2Settings.SingleOrDefault(s => s.Year == nowdaystartTime.Year);
                    if (nowco2 != null)
                    {
                        YearCO2Total = NowYearTotalvalue * nowco2.Value;
                    }
                    decimal AfterCO2Total = 0;//去年碳排
                    var afterco2 = co2Settings.SingleOrDefault(s => s.Year == afterYearstartTime.Year);
                    if (afterco2 != null)
                    {
                        AfterCO2Total = AfterYearTotalvalue * afterco2.Value;
                    }
                    decimal LoseCo2 = 0;//減碳排
                    if (AfterCO2Total > YearCO2Total)//去年碳排大於今年碳排
                    {
                        LoseCo2 = AfterCO2Total - YearCO2Total;
                    }
                    else if (YearCO2Total > AfterCO2Total & AfterCO2Total != 0)//今年碳排大於去年碳排 且 去年碳排不等於0
                    {
                        LoseCo2 = YearCO2Total - AfterCO2Total;
                    }
                    return new
                    {
                        NowDayTotal = NowDayTotalvalue,
                        AfterDayTotal = AfterDayTotalvalue,
                        NowMonthTotal = NowMonthTotalvalue,
                        AfterMonthTotal = AfterMonthTotalvalue,
                        YearCO2Total = YearCO2Total,
                        LoseCo2 = LoseCo2,
                        Temp = 0,
                        Percent = Percent,
                        RSV = ElectricWeb.Vab,
                        STV = ElectricWeb.Vbc,
                        TRV = ElectricWeb.Vca,
                        RA = ElectricWeb.Ia,
                        SA = ElectricWeb.Ib,
                        TA = ElectricWeb.Ic,
                        kW = ElectricWeb.Kwtot,
                        kVA = ElectricWeb.KVARtot,
                        PF = ElectricWeb.PFavg,
                        kVAR = ElectricWeb.KVARtot
                    };
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return new();
            }
        }
        /// <summary>
        /// 驟降驟升事件
        /// </summary>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <param name="DeviceNumber">設備編號</param>
        /// <returns></returns>
        public object Search_Sag_Swell_Info(string CardNumber, string BordNumber, int DeviceNumber)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = $"USE [EwatchIntegrate_Log] SELECT TOP(20) ttimen,Sag_Swell_Type,Cycle,Data,Phase,Sag_Swell_Begin,Sag_Swell_End FROM Electric_Sag_Swell_{CardNumber}{BordNumber} WHERE DeviceNumber = {DeviceNumber} ORDER BY ttime DESC";
                    List<object> _Logs = new List<object>();
                    _Logs = connection.Query<object>(sql).ToList();
                    List<string> TitleName = new List<string>();
                    TitleName.Add("時間");
                    TitleName.Add("驟降驟升類型");
                    TitleName.Add("電壓驟降驟升之持續 cycle 數");
                    TitleName.Add("電壓驟降驟升發生時百分比");
                    TitleName.Add("相電壓");
                    TitleName.Add("開始發生日期與時間");
                    TitleName.Add("結束日期與時間");
                    return new
                    {
                        TitleName = TitleName,
                        data = _Logs
                    };
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return new();
            }
        }
        /// <summary>
        /// 超約告警事件
        /// </summary>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <param name="DeviceNumber">設備編號</param>
        /// <returns></returns>
        public object Search_Alarm_Info(string CardNumber, string BordNumber, int DeviceNumber)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = $"USE [EwatchIntegrate_Log] SELECT TOP(20) ttimen,AlarmItem,AlarmData,AlarmTime FROM ElectricAlarm_{CardNumber}{BordNumber} WHERE DeviceNumber = {DeviceNumber} ORDER BY ttime DESC";
                    List<object> _Logs = new List<object>();
                    _Logs = connection.Query<object>(sql).ToList();
                    List<string> TitleName = new List<string>();
                    TitleName.Add("時間");
                    TitleName.Add("警報項目");
                    TitleName.Add("警報發生或解除時之百分比");
                    TitleName.Add("發生日期與時間");
                    return new
                    {
                        TitleName = TitleName,
                        data = _Logs
                    };
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return new();
            }
        }
        #endregion
        #region 曲線圖查詢
        /// <summary>
        /// 報表曲線圖查詢
        /// </summary>
        /// <param name="StartTime">開始時間</param>
        /// <param name="EndTime">結束時間</param>
        /// <param name="deviceSetting">設備資訊</param>
        /// <param name="Kind">種類</param>
        /// <param name="Item">項次</param>
        /// <returns></returns>
        public object Search_Chart(DateTime StartTime, DateTime EndTime, DeviceSetting deviceSetting, string Kind, int Item)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = $"USE [EwatchIntegrate_Log] SELECT  * FROM Electric_{deviceSetting.CardNumber}{deviceSetting.BordNumber} WHERE DeviceNumber = {deviceSetting.DeviceNumber} AND (ttime >= '{StartTime:yyyyMMdd000000}' AND ttime <= '{EndTime:yyyyMMdd235959}')";
                    List<DateTime> XAxis = new List<DateTime>();
                    List<object> Series = new List<object>();
                    var data = connection.Query<ElectricInformation_Log>(sql);
                    XAxis = data.Select(t => t.ttimen).ToList();
                    switch (Kind)
                    {
                        case "三相電壓":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "R相電壓",
                                                data = data.Select(s => s.Va).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "S相電壓",
                                                data = data.Select(s => s.Vb).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "T相電壓",
                                                data = data.Select(s => s.Vc).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "平均相電壓",
                                                data = data.Select(s => s.Vnavg).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 1:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "R相電壓",
                                                data = data.Select(s => s.Va).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 2:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "S相電壓",
                                                data = data.Select(s => s.Vb).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 3:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "T相電壓",
                                                data = data.Select(s => s.Vc).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                }
                            }
                            break;
                        case "三相電流":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "R相電流",
                                                data = data.Select(s => s.Ia).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "S相電流",
                                                data = data.Select(s => s.Ib).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "T相電流",
                                                data = data.Select(s => s.Ic).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "平均相電流",
                                                data = data.Select(s => s.Iavg).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 1:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "R相電流",
                                                data = data.Select(s => s.Ia).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 2:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "S相電流",
                                                data = data.Select(s => s.Ib).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 3:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "T相電流",
                                                data = data.Select(s => s.Ic).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                }
                            }
                            break;
                        case "即時需量":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "R相即時虛量",
                                                data = data.Select(s => s.Kwa).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "S相即時虛量",
                                                data = data.Select(s => s.Kwb).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "T相即時虛量",
                                                data = data.Select(s => s.Kwc).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "總即時虛量",
                                                data = data.Select(s => s.Kwtot).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 1:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "R相即時虛量",
                                                data = data.Select(s => s.Kwa).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 2:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "S相即時虛量",
                                                data = data.Select(s => s.Kwb).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 3:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "T相即時虛量",
                                                data = data.Select(s => s.Kwc).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                }
                            }
                            break;
                        case "虛功率":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "R相虛功率",
                                                data = data.Select(s => s.KVARa).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "S相虛功率",
                                                data = data.Select(s => s.KVARb).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "T相虛功率",
                                                data = data.Select(s => s.KVARc).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "總虛功率",
                                                data = data.Select(s => s.KVARtot).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 1:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "R相虛功率",
                                                data = data.Select(s => s.KVARa).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 2:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "S相虛功率",
                                                data = data.Select(s => s.KVARb).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 3:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "T相虛功率",
                                                data = data.Select(s => s.KVARc).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                }
                            }
                            break;
                        case "視在功率":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "R相視在功率",
                                                data = data.Select(s => s.KVAa).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "S相視在功率",
                                                data = data.Select(s => s.KVAb).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "T相視在功率",
                                                data = data.Select(s => s.KVAc).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "總視在功率",
                                                data = data.Select(s => s.KVAtot).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 1:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "R相視在功率",
                                                data = data.Select(s => s.KVAa).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 2:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "S相視在功率",
                                                data = data.Select(s => s.KVAb).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 3:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "T相視在功率",
                                                data = data.Select(s => s.KVAc).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                }
                            }
                            break;
                        case "功率因數":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "R相功率因數",
                                                data = data.Select(s => s.PFa).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "S相功率因數",
                                                data = data.Select(s => s.PFb).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "T相功率因數",
                                                data = data.Select(s => s.PFc).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "總功率因數",
                                                data = data.Select(s => s.PFavg).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 1:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "R相功率因數",
                                                data = data.Select(s => s.PFa).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 2:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "S相功率因數",
                                                data = data.Select(s => s.PFb).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 3:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "T相功率因數",
                                                data = data.Select(s => s.PFc).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                }
                            }
                            break;
                        case "電壓相位角":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        {
                                            Series.Add(new
                                            {
                                                Name = "R相電壓相位角",
                                                data = data.Select(s => s.PhaseAngleVa).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "S相電壓相位角",
                                                data = data.Select(s => s.PhaseAngleVb).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = "T相電壓相位角",
                                                data = data.Select(s => s.PhaseAngleVc).ToList(),
                                                Type = "line"
                                            });
                                        }
                                        break;
                                    case 1:
                                        Series.Add(new
                                        {
                                            Name = "R相電壓相位角",
                                            data = data.Select(s => s.PhaseAngleVa).ToList(),
                                            Type = "line"
                                        });
                                        break;
                                    case 2:
                                        Series.Add(new
                                        {
                                            Name = "S相電壓相位角",
                                            data = data.Select(s => s.PhaseAngleVb).ToList(),
                                            Type = "line"
                                        });
                                        break;
                                    case 3:
                                        Series.Add(new
                                        {
                                            Name = "T相電壓相位角",
                                            data = data.Select(s => s.PhaseAngleVc).ToList(),
                                            Type = "line"
                                        });
                                        break;
                                }
                            }
                            break;
                        case "電流相位角":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        Series.Add(new
                                        {
                                            Name = "R相電流相位角",
                                            data = data.Select(s => s.PhaseAngleIa).ToList(),
                                            Type = "line"
                                        });
                                        Series.Add(new
                                        {
                                            Name = "S相電流相位角",
                                            data = data.Select(s => s.PhaseAngleIb).ToList(),
                                            Type = "line"
                                        });
                                        Series.Add(new
                                        {
                                            Name = "T相電流相位角",
                                            data = data.Select(s => s.PhaseAngleIc).ToList(),
                                            Type = "line"
                                        });
                                        break;
                                    case 1:
                                        Series.Add(new
                                        {
                                            Name = "R相電流相位角",
                                            data = data.Select(s => s.PhaseAngleIa).ToList(),
                                            Type = "line"
                                        });
                                        break;
                                    case 2:
                                        Series.Add(new
                                        {
                                            Name = "S相電流相位角",
                                            data = data.Select(s => s.PhaseAngleIb).ToList(),
                                            Type = "line"
                                        });
                                        break;
                                    case 3:
                                        Series.Add(new
                                        {
                                            Name = "T相電流相位角",
                                            data = data.Select(s => s.PhaseAngleIc).ToList(),
                                            Type = "line"
                                        });
                                        break;
                                }
                            }
                            break;
                        case "三線電壓":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        Series.Add(new
                                        {
                                            Name = "R線電壓",
                                            data = data.Select(s => s.Vab).ToList(),
                                            Type = "line"
                                        });
                                        Series.Add(new
                                        {
                                            Name = "S線電壓",
                                            data = data.Select(s => s.Vbc).ToList(),
                                            Type = "line"
                                        });
                                        Series.Add(new
                                        {
                                            Name = "T線電壓",
                                            data = data.Select(s => s.Vca).ToList(),
                                            Type = "line"
                                        });
                                        Series.Add(new
                                        {
                                            Name = "平均線電壓",
                                            data = data.Select(s => s.Vlavg).ToList(),
                                            Type = "line"
                                        });
                                        break;
                                    case 1:
                                        Series.Add(new
                                        {
                                            Name = "RS線電壓",
                                            data = data.Select(s => s.Vab).ToList(),
                                            Type = "line"
                                        });
                                        break;
                                    case 2:
                                        Series.Add(new
                                        {
                                            Name = "ST線電壓",
                                            data = data.Select(s => s.Vbc).ToList(),
                                            Type = "line"
                                        });
                                        break;
                                    case 3:
                                        Series.Add(new
                                        {
                                            Name = "TR線電壓",
                                            data = data.Select(s => s.Vca).ToList(),
                                            Type = "line"
                                        });
                                        break;
                                }
                            }
                            break;
                        case "頻率":
                            {
                                Series.Add(new
                                {
                                    Name = "頻率",
                                    data = data.Select(s => s.Freq).ToList(),
                                    Type = "line"
                                });
                                break;
                            }
                        case "累積功率":
                            {
                                XAxis = new List<DateTime>();
                                string sqltot = $"USE [EwatchIntegrateTotal] SELECT *  FROM Electric_Log WHERE CardNumber = '{deviceSetting.CardNumber}' AND BordNumber = '{deviceSetting.BordNumber}' AND DeviceNumber = {deviceSetting.DeviceNumber} AND (ttime >= '{StartTime:yyyyMMdd000000}' AND ttime <= '{EndTime:yyyyMMdd235959}')";
                                var datatot = connection.Query<ElectricTotal_Log>(sql);
                                XAxis = datatot.Select(t => t.ttimen).ToList();
                                Series.Add(new
                                {
                                    Name = "累積功率",
                                    data = datatot.Select(s => s.Total).ToList(),
                                    Type = "bar"
                                });
                                break;
                            }
                        case "R相電壓諧波":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-Reserved",
                                                data = data.Select(s => s.RV_HD).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-1",
                                                data = data.Select(s => s.RV_HD1).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-2",
                                                data = data.Select(s => s.RV_HD2).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-3",
                                                data = data.Select(s => s.RV_HD3).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-4",
                                                data = data.Select(s => s.RV_HD4).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-5",
                                                data = data.Select(s => s.RV_HD5).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-6",
                                                data = data.Select(s => s.RV_HD6).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-7",
                                                data = data.Select(s => s.RV_HD7).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-8",
                                                data = data.Select(s => s.RV_HD8).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-9",
                                                data = data.Select(s => s.RV_HD9).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-10",
                                                data = data.Select(s => s.RV_HD10).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-11",
                                                data = data.Select(s => s.RV_HD11).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-12",
                                                data = data.Select(s => s.RV_HD12).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-13",
                                                data = data.Select(s => s.RV_HD13).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-14",
                                                data = data.Select(s => s.RV_HD14).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-15",
                                                data = data.Select(s => s.RV_HD15).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-16",
                                                data = data.Select(s => s.RV_HD16).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-17",
                                                data = data.Select(s => s.RV_HD17).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-18",
                                                data = data.Select(s => s.RV_HD18).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-19",
                                                data = data.Select(s => s.RV_HD19).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-20",
                                                data = data.Select(s => s.RV_HD20).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-21",
                                                data = data.Select(s => s.RV_HD21).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-22",
                                                data = data.Select(s => s.RV_HD22).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-23",
                                                data = data.Select(s => s.RV_HD23).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-24",
                                                data = data.Select(s => s.RV_HD24).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-25",
                                                data = data.Select(s => s.RV_HD25).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-26",
                                                data = data.Select(s => s.RV_HD26).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-27",
                                                data = data.Select(s => s.RV_HD27).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-28",
                                                data = data.Select(s => s.RV_HD28).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-29",
                                                data = data.Select(s => s.RV_HD29).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-30",
                                                data = data.Select(s => s.RV_HD30).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-31",
                                                data = data.Select(s => s.RV_HD31).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 1:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD1).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 2:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD2).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 3:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD3).ToList(),
                                                Type = "line"
                                            });
                                            return new
                                            {
                                                XAxis = XAxis,
                                                Series = Series
                                            };
                                        }
                                    case 4:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD4).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 5:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD5).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 6:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD6).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 7:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD7).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 8:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD8).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 9:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD9).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 10:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD10).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 11:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD11).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 12:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD12).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 13:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD13).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 14:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD14).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 15:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD15).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 16:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD16).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 17:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD17).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 18:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD18).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 19:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD19).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 20:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD20).ToList(),
                                                Type = "line"
                                            });
                                            return new
                                            {
                                                XAxis = XAxis,
                                                Series = Series
                                            };
                                        }
                                    case 21:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD21).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 22:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD22).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 23:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD23).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 24:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD24).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 25:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD25).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 26:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD26).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 27:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD27).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 28:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD28).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 29:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD29).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 30:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD30).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 31:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電壓諧波-{Item}",
                                                data = data.Select(s => s.RV_HD31).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                }
                            }
                            break;
                        case "S相電壓諧波":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-Reserved",
                                                data = data.Select(s => s.SV_HD).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-1",
                                                data = data.Select(s => s.SV_HD1).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-2",
                                                data = data.Select(s => s.SV_HD2).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-3",
                                                data = data.Select(s => s.SV_HD3).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-4",
                                                data = data.Select(s => s.SV_HD4).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-5",
                                                data = data.Select(s => s.SV_HD5).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-6",
                                                data = data.Select(s => s.SV_HD6).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-7",
                                                data = data.Select(s => s.SV_HD7).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-8",
                                                data = data.Select(s => s.SV_HD8).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-9",
                                                data = data.Select(s => s.SV_HD9).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-10",
                                                data = data.Select(s => s.SV_HD10).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-11",
                                                data = data.Select(s => s.SV_HD11).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-12",
                                                data = data.Select(s => s.SV_HD12).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-13",
                                                data = data.Select(s => s.SV_HD13).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-14",
                                                data = data.Select(s => s.SV_HD14).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-15",
                                                data = data.Select(s => s.SV_HD15).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-16",
                                                data = data.Select(s => s.SV_HD16).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-17",
                                                data = data.Select(s => s.SV_HD17).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-18",
                                                data = data.Select(s => s.SV_HD18).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-19",
                                                data = data.Select(s => s.SV_HD19).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-20",
                                                data = data.Select(s => s.SV_HD20).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-21",
                                                data = data.Select(s => s.SV_HD21).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-22",
                                                data = data.Select(s => s.SV_HD22).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-23",
                                                data = data.Select(s => s.SV_HD23).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-24",
                                                data = data.Select(s => s.SV_HD24).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-25",
                                                data = data.Select(s => s.SV_HD25).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-26",
                                                data = data.Select(s => s.SV_HD26).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-27",
                                                data = data.Select(s => s.SV_HD27).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-28",
                                                data = data.Select(s => s.SV_HD28).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-29",
                                                data = data.Select(s => s.SV_HD29).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-30",
                                                data = data.Select(s => s.SV_HD30).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-31",
                                                data = data.Select(s => s.SV_HD31).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 1:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD1).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 2:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD2).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 3:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD3).ToList(),
                                                Type = "line"
                                            });
                                            return new
                                            {
                                                XAxis = XAxis,
                                                Series = Series
                                            };
                                        }
                                    case 4:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD4).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 5:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD5).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 6:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD6).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 7:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD7).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 8:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD8).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 9:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD9).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 10:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD10).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 11:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD11).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 12:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD12).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 13:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD13).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 14:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD14).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 15:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD15).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 16:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD16).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 17:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD17).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 18:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD18).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 19:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD19).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 20:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD20).ToList(),
                                                Type = "line"
                                            });
                                            return new
                                            {
                                                XAxis = XAxis,
                                                Series = Series
                                            };
                                        }
                                    case 21:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD21).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 22:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD22).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 23:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD23).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 24:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD24).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 25:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD25).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 26:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD26).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 27:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD27).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 28:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD28).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 29:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD29).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 30:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD30).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 31:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電壓諧波-{Item}",
                                                data = data.Select(s => s.SV_HD31).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                }
                            }
                            break;
                        case "T相電壓諧波":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-Teserved",
                                                data = data.Select(s => s.TV_HD).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-1",
                                                data = data.Select(s => s.TV_HD1).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-2",
                                                data = data.Select(s => s.TV_HD2).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-3",
                                                data = data.Select(s => s.TV_HD3).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-4",
                                                data = data.Select(s => s.TV_HD4).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-5",
                                                data = data.Select(s => s.TV_HD5).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-6",
                                                data = data.Select(s => s.TV_HD6).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-7",
                                                data = data.Select(s => s.TV_HD7).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-8",
                                                data = data.Select(s => s.TV_HD8).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-9",
                                                data = data.Select(s => s.TV_HD9).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-10",
                                                data = data.Select(s => s.TV_HD10).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-11",
                                                data = data.Select(s => s.TV_HD11).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-12",
                                                data = data.Select(s => s.TV_HD12).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-13",
                                                data = data.Select(s => s.TV_HD13).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-14",
                                                data = data.Select(s => s.TV_HD14).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-15",
                                                data = data.Select(s => s.TV_HD15).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-16",
                                                data = data.Select(s => s.TV_HD16).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-17",
                                                data = data.Select(s => s.TV_HD17).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-18",
                                                data = data.Select(s => s.TV_HD18).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-19",
                                                data = data.Select(s => s.TV_HD19).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-20",
                                                data = data.Select(s => s.TV_HD20).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-21",
                                                data = data.Select(s => s.TV_HD21).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-22",
                                                data = data.Select(s => s.TV_HD22).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-23",
                                                data = data.Select(s => s.TV_HD23).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-24",
                                                data = data.Select(s => s.TV_HD24).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-25",
                                                data = data.Select(s => s.TV_HD25).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-26",
                                                data = data.Select(s => s.TV_HD26).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-27",
                                                data = data.Select(s => s.TV_HD27).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-28",
                                                data = data.Select(s => s.TV_HD28).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-29",
                                                data = data.Select(s => s.TV_HD29).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-30",
                                                data = data.Select(s => s.TV_HD30).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-31",
                                                data = data.Select(s => s.TV_HD31).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 1:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD1).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 2:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD2).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 3:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD3).ToList(),
                                                Type = "line"
                                            });
                                            return new
                                            {
                                                XAxis = XAxis,
                                                Series = Series
                                            };
                                        }
                                    case 4:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD4).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 5:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD5).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 6:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD6).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 7:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD7).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 8:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD8).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 9:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD9).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 10:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD10).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 11:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD11).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 12:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD12).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 13:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD13).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 14:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD14).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 15:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD15).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 16:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD16).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 17:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD17).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 18:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD18).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 19:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD19).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 20:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD20).ToList(),
                                                Type = "line"
                                            });
                                            return new
                                            {
                                                XAxis = XAxis,
                                                Series = Series
                                            };
                                        }
                                    case 21:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD21).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 22:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD22).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 23:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD23).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 24:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD24).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 25:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD25).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 26:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD26).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 27:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD27).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 28:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD28).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 29:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD29).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 30:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD30).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 31:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電壓諧波-{Item}",
                                                data = data.Select(s => s.TV_HD31).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                }
                            }
                            break;
                        case "R相電流諧波":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-Reserved",
                                                data = data.Select(s => s.RI_HD).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-1",
                                                data = data.Select(s => s.RI_HD1).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-2",
                                                data = data.Select(s => s.RI_HD2).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-3",
                                                data = data.Select(s => s.RI_HD3).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-4",
                                                data = data.Select(s => s.RI_HD4).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-5",
                                                data = data.Select(s => s.RI_HD5).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-6",
                                                data = data.Select(s => s.RI_HD6).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-7",
                                                data = data.Select(s => s.RI_HD7).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-8",
                                                data = data.Select(s => s.RI_HD8).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-9",
                                                data = data.Select(s => s.RI_HD9).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-10",
                                                data = data.Select(s => s.RI_HD10).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-11",
                                                data = data.Select(s => s.RI_HD11).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-12",
                                                data = data.Select(s => s.RI_HD12).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-13",
                                                data = data.Select(s => s.RI_HD13).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-14",
                                                data = data.Select(s => s.RI_HD14).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-15",
                                                data = data.Select(s => s.RI_HD15).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-16",
                                                data = data.Select(s => s.RI_HD16).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-17",
                                                data = data.Select(s => s.RI_HD17).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-18",
                                                data = data.Select(s => s.RI_HD18).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-19",
                                                data = data.Select(s => s.RI_HD19).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-20",
                                                data = data.Select(s => s.RI_HD20).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-21",
                                                data = data.Select(s => s.RI_HD21).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-22",
                                                data = data.Select(s => s.RI_HD22).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-23",
                                                data = data.Select(s => s.RI_HD23).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-24",
                                                data = data.Select(s => s.RI_HD24).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-25",
                                                data = data.Select(s => s.RI_HD25).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-26",
                                                data = data.Select(s => s.RI_HD26).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-27",
                                                data = data.Select(s => s.RI_HD27).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-28",
                                                data = data.Select(s => s.RI_HD28).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-29",
                                                data = data.Select(s => s.RI_HD29).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-30",
                                                data = data.Select(s => s.RI_HD30).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-31",
                                                data = data.Select(s => s.RI_HD31).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 1:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD1).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 2:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD2).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 3:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD3).ToList(),
                                                Type = "line"
                                            });
                                            return new
                                            {
                                                XAxis = XAxis,
                                                Series = Series
                                            };
                                        }
                                    case 4:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD4).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 5:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD5).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 6:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD6).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 7:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD7).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 8:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD8).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 9:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD9).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 10:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD10).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 11:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD11).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 12:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD12).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 13:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD13).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 14:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD14).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 15:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD15).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 16:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD16).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 17:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD17).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 18:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD18).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 19:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD19).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 20:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD20).ToList(),
                                                Type = "line"
                                            });
                                            return new
                                            {
                                                XAxis = XAxis,
                                                Series = Series
                                            };
                                        }
                                    case 21:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD21).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 22:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD22).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 23:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD23).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 24:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD24).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 25:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD25).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 26:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD26).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 27:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD27).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 28:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD28).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 29:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD29).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 30:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD30).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 31:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"R相電流諧波-{Item}",
                                                data = data.Select(s => s.RI_HD31).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                }
                            }
                            break;
                        case "S相電流諧波":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-Reserved",
                                                data = data.Select(s => s.SI_HD).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-1",
                                                data = data.Select(s => s.SI_HD1).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-2",
                                                data = data.Select(s => s.SI_HD2).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-3",
                                                data = data.Select(s => s.SI_HD3).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-4",
                                                data = data.Select(s => s.SI_HD4).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-5",
                                                data = data.Select(s => s.SI_HD5).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-6",
                                                data = data.Select(s => s.SI_HD6).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-7",
                                                data = data.Select(s => s.SI_HD7).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-8",
                                                data = data.Select(s => s.SI_HD8).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-9",
                                                data = data.Select(s => s.SI_HD9).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-10",
                                                data = data.Select(s => s.SI_HD10).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-11",
                                                data = data.Select(s => s.SI_HD11).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-12",
                                                data = data.Select(s => s.SI_HD12).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-13",
                                                data = data.Select(s => s.SI_HD13).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-14",
                                                data = data.Select(s => s.SI_HD14).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-15",
                                                data = data.Select(s => s.SI_HD15).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-16",
                                                data = data.Select(s => s.SI_HD16).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-17",
                                                data = data.Select(s => s.SI_HD17).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-18",
                                                data = data.Select(s => s.SI_HD18).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-19",
                                                data = data.Select(s => s.SI_HD19).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-20",
                                                data = data.Select(s => s.SI_HD20).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-21",
                                                data = data.Select(s => s.SI_HD21).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-22",
                                                data = data.Select(s => s.SI_HD22).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-23",
                                                data = data.Select(s => s.SI_HD23).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-24",
                                                data = data.Select(s => s.SI_HD24).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-25",
                                                data = data.Select(s => s.SI_HD25).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-26",
                                                data = data.Select(s => s.SI_HD26).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-27",
                                                data = data.Select(s => s.SI_HD27).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-28",
                                                data = data.Select(s => s.SI_HD28).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-29",
                                                data = data.Select(s => s.SI_HD29).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-30",
                                                data = data.Select(s => s.SI_HD30).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-31",
                                                data = data.Select(s => s.SI_HD31).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 1:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD1).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 2:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD2).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 3:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD3).ToList(),
                                                Type = "line"
                                            });
                                            return new
                                            {
                                                XAxis = XAxis,
                                                Series = Series
                                            };
                                        }
                                    case 4:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD4).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 5:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD5).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 6:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD6).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 7:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD7).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 8:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD8).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 9:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD9).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 10:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD10).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 11:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD11).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 12:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD12).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 13:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD13).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 14:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD14).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 15:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD15).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 16:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD16).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 17:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD17).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 18:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD18).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 19:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD19).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 20:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD20).ToList(),
                                                Type = "line"
                                            });
                                            return new
                                            {
                                                XAxis = XAxis,
                                                Series = Series
                                            };
                                        }
                                    case 21:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD21).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 22:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD22).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 23:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD23).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 24:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD24).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 25:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD25).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 26:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD26).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 27:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD27).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 28:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD28).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 29:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD29).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 30:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD30).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 31:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"S相電流諧波-{Item}",
                                                data = data.Select(s => s.SI_HD31).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                }
                            }
                            break;
                        case "T相電流諧波":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-Teserved",
                                                data = data.Select(s => s.TI_HD).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-1",
                                                data = data.Select(s => s.TI_HD1).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-2",
                                                data = data.Select(s => s.TI_HD2).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-3",
                                                data = data.Select(s => s.TI_HD3).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-4",
                                                data = data.Select(s => s.TI_HD4).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-5",
                                                data = data.Select(s => s.TI_HD5).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-6",
                                                data = data.Select(s => s.TI_HD6).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-7",
                                                data = data.Select(s => s.TI_HD7).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-8",
                                                data = data.Select(s => s.TI_HD8).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-9",
                                                data = data.Select(s => s.TI_HD9).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-10",
                                                data = data.Select(s => s.TI_HD10).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-11",
                                                data = data.Select(s => s.TI_HD11).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-12",
                                                data = data.Select(s => s.TI_HD12).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-13",
                                                data = data.Select(s => s.TI_HD13).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-14",
                                                data = data.Select(s => s.TI_HD14).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-15",
                                                data = data.Select(s => s.TI_HD15).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-16",
                                                data = data.Select(s => s.TI_HD16).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-17",
                                                data = data.Select(s => s.TI_HD17).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-18",
                                                data = data.Select(s => s.TI_HD18).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-19",
                                                data = data.Select(s => s.TI_HD19).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-20",
                                                data = data.Select(s => s.TI_HD20).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-21",
                                                data = data.Select(s => s.TI_HD21).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-22",
                                                data = data.Select(s => s.TI_HD22).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-23",
                                                data = data.Select(s => s.TI_HD23).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-24",
                                                data = data.Select(s => s.TI_HD24).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-25",
                                                data = data.Select(s => s.TI_HD25).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-26",
                                                data = data.Select(s => s.TI_HD26).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-27",
                                                data = data.Select(s => s.TI_HD27).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-28",
                                                data = data.Select(s => s.TI_HD28).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-29",
                                                data = data.Select(s => s.TI_HD29).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-30",
                                                data = data.Select(s => s.TI_HD30).ToList(),
                                                Type = "line"
                                            });
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-31",
                                                data = data.Select(s => s.TI_HD31).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 1:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD1).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 2:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD2).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 3:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD3).ToList(),
                                                Type = "line"
                                            });
                                            return new
                                            {
                                                XAxis = XAxis,
                                                Series = Series
                                            };
                                        }
                                    case 4:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD4).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 5:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD5).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 6:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD6).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 7:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD7).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 8:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD8).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 9:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD9).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 10:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD10).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 11:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD11).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 12:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD12).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 13:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD13).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 14:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD14).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 15:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD15).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 16:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD16).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 17:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD17).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 18:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD18).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 19:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD19).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 20:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD20).ToList(),
                                                Type = "line"
                                            });
                                            return new
                                            {
                                                XAxis = XAxis,
                                                Series = Series
                                            };
                                        }
                                    case 21:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD21).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 22:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD22).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 23:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD23).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 24:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD24).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 25:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD25).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 26:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD26).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 27:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD27).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 28:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD28).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 29:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD29).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 30:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD30).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                    case 31:
                                        {
                                            Series.Add(new
                                            {
                                                Name = $"T相電流諧波-{Item}",
                                                data = data.Select(s => s.TI_HD31).ToList(),
                                                Type = "line"
                                            });
                                            break;
                                        }
                                }
                            }
                            break;
                    }
                    return new
                    {
                        XAxis = XAxis,
                        Series = Series
                    };
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return new();
            }
        }
        /// <summary>
        /// 下載報表資訊
        /// </summary>
        /// <param name="StartTime">開始時間</param>
        /// <param name="EndTime">結束時間</param>
        /// <param name="deviceSetting">設備資訊</param>
        /// <param name="Kind">種類</param>
        /// <param name="Item">項次</param>
        /// <returns></returns>
        public FileStreamResult? DownLoad_Device_Chart_Report(DateTime StartTime, DateTime EndTime, DeviceSetting deviceSetting, string Kind, int Item)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string value = "ttimen";
                    List<string> TitleName = new List<string>();
                    TitleName.Add("時間");
                    switch (Kind)
                    {
                        case "三相電壓":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        value += ",Va,Vb,Vc,Vnavg";
                                        TitleName.Add("R相電壓");
                                        TitleName.Add("S相電壓");
                                        TitleName.Add("T相電壓");
                                        TitleName.Add("平均相電壓");
                                        break;
                                    case 1:
                                        value += ",Va";
                                        TitleName.Add("R相電壓");
                                        break;
                                    case 2:
                                        value += ",Vb";
                                        TitleName.Add("S相電壓");
                                        break;
                                    case 3:
                                        value += ",Vc";
                                        TitleName.Add("T相電壓");
                                        break;
                                }
                            }
                            break;
                        case "三相電流":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        value += ",Ia,Ib,Ic,Iavg";
                                        TitleName.Add("R相電流");
                                        TitleName.Add("S相電流");
                                        TitleName.Add("T相電流");
                                        TitleName.Add("平均相電流");
                                        break;
                                    case 1:
                                        value += ",Ia";
                                        TitleName.Add("R相電流");
                                        break;
                                    case 2:
                                        value += ",Ib";
                                        TitleName.Add("S相電流");
                                        break;
                                    case 3:
                                        value += ",Ic";
                                        TitleName.Add("T相電流");
                                        break;
                                }
                            }
                            break;
                        case "即時需量":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        value += ",Kwa,Kwb,Kwc,Kwtot";
                                        TitleName.Add("R相即時需量");
                                        TitleName.Add("S相即時需量");
                                        TitleName.Add("T相即時需量");
                                        TitleName.Add("總即時需量");
                                        break;
                                    case 1:
                                        value += ",Kwa";
                                        TitleName.Add("R相即時需量");
                                        break;
                                    case 2:
                                        value += ",Kwb";
                                        TitleName.Add("S相即時需量");
                                        break;
                                    case 3:
                                        value += ",Kwc";
                                        TitleName.Add("T相即時需量");
                                        break;
                                }
                            }
                            break;
                        case "虛功率":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        value += ",KVARa,KVARb,KVARc,KVARtot";
                                        TitleName.Add("R相虛功率");
                                        TitleName.Add("S相虛功率");
                                        TitleName.Add("T相虛功率");
                                        TitleName.Add("總虛功率");
                                        break;
                                    case 1:
                                        value += ",KVARa";
                                        TitleName.Add("R相虛功率");
                                        break;
                                    case 2:
                                        value += ",KVARb";
                                        TitleName.Add("S相虛功率");
                                        break;
                                    case 3:
                                        value += ",KVARc";
                                        TitleName.Add("T相虛功率");
                                        break;
                                }
                            }
                            break;
                        case "視在功率":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        value += ",KVAa,KVAb,KVAc,KVAtot";
                                        TitleName.Add("R相視在功率");
                                        TitleName.Add("S相視在功率");
                                        TitleName.Add("T相視在功率");
                                        TitleName.Add("總視在功率");
                                        break;
                                    case 1:
                                        value += ",KVAa";
                                        TitleName.Add("R相視在功率");
                                        break;
                                    case 2:
                                        value += ",KVAb";
                                        TitleName.Add("S相視在功率");
                                        break;
                                    case 3:
                                        value += ",KVAc";
                                        TitleName.Add("T相視在功率");
                                        break;
                                }
                            }
                            break;
                        case "功率因數":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        value += ",PFa,PFb,PFc,PFavg";
                                        TitleName.Add("R相功率因數");
                                        TitleName.Add("S相功率因數");
                                        TitleName.Add("T相功率因數");
                                        TitleName.Add("總功率因數");
                                        break;
                                    case 1:
                                        value += ",PFa";
                                        TitleName.Add("R相功率因數");
                                        break;
                                    case 2:
                                        value += ",PFb";
                                        TitleName.Add("S相功率因數");
                                        break;
                                    case 3:
                                        value += ",PFc";
                                        TitleName.Add("T相功率因數");
                                        break;
                                }
                            }
                            break;
                        case "電壓相位角":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        value += ",PhaseAngleVa,PhaseAngleVb,PhaseAngleVc";
                                        TitleName.Add("R相電壓相位角");
                                        TitleName.Add("S相電壓相位角");
                                        TitleName.Add("T相電壓相位角");
                                        break;
                                    case 1:
                                        value += ",PhaseAngleVa";
                                        TitleName.Add("R相電壓相位角");
                                        break;
                                    case 2:
                                        value += ",PhaseAngleVb";
                                        TitleName.Add("S相電壓相位角");
                                        break;
                                    case 3:
                                        value += ",PhaseAngleVc";
                                        TitleName.Add("T相電壓相位角");
                                        break;
                                }
                            }
                            break;
                        case "電流相位角":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        value += ",PhaseAngleIa,PhaseAngleIb,PhaseAngleIc";
                                        TitleName.Add("R相電流相位角");
                                        TitleName.Add("S相電流相位角");
                                        TitleName.Add("T相電流相位角");
                                        break;
                                    case 1:
                                        value += ",PhaseAngleIa";
                                        TitleName.Add("R相電流相位角");
                                        break;
                                    case 2:
                                        value += ",PhaseAngleIb";
                                        TitleName.Add("S相電流相位角");
                                        break;
                                    case 3:
                                        value += ",PhaseAngleIc";
                                        TitleName.Add("T相電流相位角");
                                        break;
                                }
                            }
                            break;
                        case "三線電壓":
                            {
                                switch (Item)
                                {
                                    case 0:
                                        value += ",Vab,Vbc,Vca,Vlavg";
                                        TitleName.Add("RS線電壓");
                                        TitleName.Add("ST線電壓");
                                        TitleName.Add("TR線電壓");
                                        TitleName.Add("平均線電壓");
                                        break;
                                    case 1:
                                        value += ",Vab";
                                        TitleName.Add("RS線電壓");
                                        break;
                                    case 2:
                                        value += ",Vbc";
                                        TitleName.Add("ST線電壓");
                                        break;
                                    case 3:
                                        value += ",Vca";
                                        TitleName.Add("TR線電壓");
                                        break;
                                }
                            }
                            break;
                        case "頻率":
                            {
                                value += ",Freq";
                                TitleName.Add("頻率");
                            }
                            break;
                        case "累積功率":
                            {
                                value += ",KWH";
                                TitleName.Add("累積功率");
                            }
                            break;
                        case "R相電壓諧波":
                            {
                                value += ",RV_HD" +
                                    ",RV_HD1" +
                                    ",RV_HD2" +
                                    ",RV_HD3" +
                                    ",RV_HD4" +
                                    ",RV_HD5" +
                                    ",RV_HD6" +
                                    ",RV_HD7" +
                                    ",RV_HD8" +
                                    ",RV_HD9" +
                                    ",RV_HD10" +
                                    ",RV_HD11" +
                                    ",RV_HD12" +
                                    ",RV_HD13" +
                                    ",RV_HD14" +
                                    ",RV_HD15" +
                                    ",RV_HD16" +
                                    ",RV_HD17" +
                                    ",RV_HD18" +
                                    ",RV_HD19" +
                                    ",RV_HD20" +
                                    ",RV_HD21" +
                                    ",RV_HD22" +
                                    ",RV_HD23" +
                                    ",RV_HD24" +
                                    ",RV_HD25" +
                                    ",RV_HD26" +
                                    ",RV_HD27" +
                                    ",RV_HD28" +
                                    ",RV_HD29" +
                                    ",RV_HD30" +
                                    ",RV_HD31";
                                TitleName.Add("R相電壓諧波-Reserved");
                                TitleName.Add("R相電壓諧波-1");
                                TitleName.Add("R相電壓諧波-2");
                                TitleName.Add("R相電壓諧波-3");
                                TitleName.Add("R相電壓諧波-4");
                                TitleName.Add("R相電壓諧波-5");
                                TitleName.Add("R相電壓諧波-6");
                                TitleName.Add("R相電壓諧波-7");
                                TitleName.Add("R相電壓諧波-8");
                                TitleName.Add("R相電壓諧波-9");
                                TitleName.Add("R相電壓諧波-10");
                                TitleName.Add("R相電壓諧波-11");
                                TitleName.Add("R相電壓諧波-12");
                                TitleName.Add("R相電壓諧波-13");
                                TitleName.Add("R相電壓諧波-14");
                                TitleName.Add("R相電壓諧波-15");
                                TitleName.Add("R相電壓諧波-16");
                                TitleName.Add("R相電壓諧波-17");
                                TitleName.Add("R相電壓諧波-18");
                                TitleName.Add("R相電壓諧波-19");
                                TitleName.Add("R相電壓諧波-20");
                                TitleName.Add("R相電壓諧波-21");
                                TitleName.Add("R相電壓諧波-22");
                                TitleName.Add("R相電壓諧波-23");
                                TitleName.Add("R相電壓諧波-24");
                                TitleName.Add("R相電壓諧波-25");
                                TitleName.Add("R相電壓諧波-26");
                                TitleName.Add("R相電壓諧波-27");
                                TitleName.Add("R相電壓諧波-28");
                                TitleName.Add("R相電壓諧波-29");
                                TitleName.Add("R相電壓諧波-30");
                                TitleName.Add("R相電壓諧波-31");
                            }
                            break;
                        case "S相電壓諧波":
                            {
                                value += ",SV_HD" +
                                    ",SV_HD1" +
                                    ",SV_HD2" +
                                    ",SV_HD3" +
                                    ",SV_HD4" +
                                    ",SV_HD5" +
                                    ",SV_HD6" +
                                    ",SV_HD7" +
                                    ",SV_HD8" +
                                    ",SV_HD9" +
                                    ",SV_HD10" +
                                    ",SV_HD11" +
                                    ",SV_HD12" +
                                    ",SV_HD13" +
                                    ",SV_HD14" +
                                    ",SV_HD15" +
                                    ",SV_HD16" +
                                    ",SV_HD17" +
                                    ",SV_HD18" +
                                    ",SV_HD19" +
                                    ",SV_HD20" +
                                    ",SV_HD21" +
                                    ",SV_HD22" +
                                    ",SV_HD23" +
                                    ",SV_HD24" +
                                    ",SV_HD25" +
                                    ",SV_HD26" +
                                    ",SV_HD27" +
                                    ",SV_HD28" +
                                    ",SV_HD29" +
                                    ",SV_HD30" +
                                    ",SV_HD31";
                                TitleName.Add("S相電壓諧波-Reserved");
                                TitleName.Add("S相電壓諧波-1");
                                TitleName.Add("S相電壓諧波-2");
                                TitleName.Add("S相電壓諧波-3");
                                TitleName.Add("S相電壓諧波-4");
                                TitleName.Add("S相電壓諧波-5");
                                TitleName.Add("S相電壓諧波-6");
                                TitleName.Add("S相電壓諧波-7");
                                TitleName.Add("S相電壓諧波-8");
                                TitleName.Add("S相電壓諧波-9");
                                TitleName.Add("S相電壓諧波-10");
                                TitleName.Add("S相電壓諧波-11");
                                TitleName.Add("S相電壓諧波-12");
                                TitleName.Add("S相電壓諧波-13");
                                TitleName.Add("S相電壓諧波-14");
                                TitleName.Add("S相電壓諧波-15");
                                TitleName.Add("S相電壓諧波-16");
                                TitleName.Add("S相電壓諧波-17");
                                TitleName.Add("S相電壓諧波-18");
                                TitleName.Add("S相電壓諧波-19");
                                TitleName.Add("S相電壓諧波-20");
                                TitleName.Add("S相電壓諧波-21");
                                TitleName.Add("S相電壓諧波-22");
                                TitleName.Add("S相電壓諧波-23");
                                TitleName.Add("S相電壓諧波-24");
                                TitleName.Add("S相電壓諧波-25");
                                TitleName.Add("S相電壓諧波-26");
                                TitleName.Add("S相電壓諧波-27");
                                TitleName.Add("S相電壓諧波-28");
                                TitleName.Add("S相電壓諧波-29");
                                TitleName.Add("S相電壓諧波-30");
                                TitleName.Add("S相電壓諧波-31");
                            }
                            break;
                        case "T相電壓諧波":
                            {
                                value += ",TV_HD" +
                                    ",TV_HD1" +
                                    ",TV_HD2" +
                                    ",TV_HD3" +
                                    ",TV_HD4" +
                                    ",TV_HD5" +
                                    ",TV_HD6" +
                                    ",TV_HD7" +
                                    ",TV_HD8" +
                                    ",TV_HD9" +
                                    ",TV_HD10" +
                                    ",TV_HD11" +
                                    ",TV_HD12" +
                                    ",TV_HD13" +
                                    ",TV_HD14" +
                                    ",TV_HD15" +
                                    ",TV_HD16" +
                                    ",TV_HD17" +
                                    ",TV_HD18" +
                                    ",TV_HD19" +
                                    ",TV_HD20" +
                                    ",TV_HD21" +
                                    ",TV_HD22" +
                                    ",TV_HD23" +
                                    ",TV_HD24" +
                                    ",TV_HD25" +
                                    ",TV_HD26" +
                                    ",TV_HD27" +
                                    ",TV_HD28" +
                                    ",TV_HD29" +
                                    ",TV_HD30" +
                                    ",TV_HD31";
                                TitleName.Add("T相電壓諧波-Reserved");
                                TitleName.Add("T相電壓諧波-1");
                                TitleName.Add("T相電壓諧波-2");
                                TitleName.Add("T相電壓諧波-3");
                                TitleName.Add("T相電壓諧波-4");
                                TitleName.Add("T相電壓諧波-5");
                                TitleName.Add("T相電壓諧波-6");
                                TitleName.Add("T相電壓諧波-7");
                                TitleName.Add("T相電壓諧波-8");
                                TitleName.Add("T相電壓諧波-9");
                                TitleName.Add("T相電壓諧波-10");
                                TitleName.Add("T相電壓諧波-11");
                                TitleName.Add("T相電壓諧波-12");
                                TitleName.Add("T相電壓諧波-13");
                                TitleName.Add("T相電壓諧波-14");
                                TitleName.Add("T相電壓諧波-15");
                                TitleName.Add("T相電壓諧波-16");
                                TitleName.Add("T相電壓諧波-17");
                                TitleName.Add("T相電壓諧波-18");
                                TitleName.Add("T相電壓諧波-19");
                                TitleName.Add("T相電壓諧波-20");
                                TitleName.Add("T相電壓諧波-21");
                                TitleName.Add("T相電壓諧波-22");
                                TitleName.Add("T相電壓諧波-23");
                                TitleName.Add("T相電壓諧波-24");
                                TitleName.Add("T相電壓諧波-25");
                                TitleName.Add("T相電壓諧波-26");
                                TitleName.Add("T相電壓諧波-27");
                                TitleName.Add("T相電壓諧波-28");
                                TitleName.Add("T相電壓諧波-29");
                                TitleName.Add("T相電壓諧波-30");
                                TitleName.Add("T相電壓諧波-31");
                            }
                            break;
                        case "R相電流諧波":
                            {
                                value += ",RI_HD" +
                                    ",RI_HD1" +
                                    ",RI_HD2" +
                                    ",RI_HD3" +
                                    ",RI_HD4" +
                                    ",RI_HD5" +
                                    ",RI_HD6" +
                                    ",RI_HD7" +
                                    ",RI_HD8" +
                                    ",RI_HD9" +
                                    ",RI_HD10" +
                                    ",RI_HD11" +
                                    ",RI_HD12" +
                                    ",RI_HD13" +
                                    ",RI_HD14" +
                                    ",RI_HD15" +
                                    ",RI_HD16" +
                                    ",RI_HD17" +
                                    ",RI_HD18" +
                                    ",RI_HD19" +
                                    ",RI_HD20" +
                                    ",RI_HD21" +
                                    ",RI_HD22" +
                                    ",RI_HD23" +
                                    ",RI_HD24" +
                                    ",RI_HD25" +
                                    ",RI_HD26" +
                                    ",RI_HD27" +
                                    ",RI_HD28" +
                                    ",RI_HD29" +
                                    ",RI_HD30" +
                                    ",RI_HD31";
                                TitleName.Add("R相電流諧波-Reserved");
                                TitleName.Add("R相電流諧波-1");
                                TitleName.Add("R相電流諧波-2");
                                TitleName.Add("R相電流諧波-3");
                                TitleName.Add("R相電流諧波-4");
                                TitleName.Add("R相電流諧波-5");
                                TitleName.Add("R相電流諧波-6");
                                TitleName.Add("R相電流諧波-7");
                                TitleName.Add("R相電流諧波-8");
                                TitleName.Add("R相電流諧波-9");
                                TitleName.Add("R相電流諧波-10");
                                TitleName.Add("R相電流諧波-11");
                                TitleName.Add("R相電流諧波-12");
                                TitleName.Add("R相電流諧波-13");
                                TitleName.Add("R相電流諧波-14");
                                TitleName.Add("R相電流諧波-15");
                                TitleName.Add("R相電流諧波-16");
                                TitleName.Add("R相電流諧波-17");
                                TitleName.Add("R相電流諧波-18");
                                TitleName.Add("R相電流諧波-19");
                                TitleName.Add("R相電流諧波-20");
                                TitleName.Add("R相電流諧波-21");
                                TitleName.Add("R相電流諧波-22");
                                TitleName.Add("R相電流諧波-23");
                                TitleName.Add("R相電流諧波-24");
                                TitleName.Add("R相電流諧波-25");
                                TitleName.Add("R相電流諧波-26");
                                TitleName.Add("R相電流諧波-27");
                                TitleName.Add("R相電流諧波-28");
                                TitleName.Add("R相電流諧波-29");
                                TitleName.Add("R相電流諧波-30");
                                TitleName.Add("R相電流諧波-31");
                            }
                            break;
                        case "S相電流諧波":
                            {
                                value += ",SI_HD" +
                                    ",SI_HD1" +
                                    ",SI_HD2" +
                                    ",SI_HD3" +
                                    ",SI_HD4" +
                                    ",SI_HD5" +
                                    ",SI_HD6" +
                                    ",SI_HD7" +
                                    ",SI_HD8" +
                                    ",SI_HD9" +
                                    ",SI_HD10" +
                                    ",SI_HD11" +
                                    ",SI_HD12" +
                                    ",SI_HD13" +
                                    ",SI_HD14" +
                                    ",SI_HD15" +
                                    ",SI_HD16" +
                                    ",SI_HD17" +
                                    ",SI_HD18" +
                                    ",SI_HD19" +
                                    ",SI_HD20" +
                                    ",SI_HD21" +
                                    ",SI_HD22" +
                                    ",SI_HD23" +
                                    ",SI_HD24" +
                                    ",SI_HD25" +
                                    ",SI_HD26" +
                                    ",SI_HD27" +
                                    ",SI_HD28" +
                                    ",SI_HD29" +
                                    ",SI_HD30" +
                                    ",SI_HD31";
                                TitleName.Add("S相電流諧波-Reserved");
                                TitleName.Add("S相電流諧波-1");
                                TitleName.Add("S相電流諧波-2");
                                TitleName.Add("S相電流諧波-3");
                                TitleName.Add("S相電流諧波-4");
                                TitleName.Add("S相電流諧波-5");
                                TitleName.Add("S相電流諧波-6");
                                TitleName.Add("S相電流諧波-7");
                                TitleName.Add("S相電流諧波-8");
                                TitleName.Add("S相電流諧波-9");
                                TitleName.Add("S相電流諧波-10");
                                TitleName.Add("S相電流諧波-11");
                                TitleName.Add("S相電流諧波-12");
                                TitleName.Add("S相電流諧波-13");
                                TitleName.Add("S相電流諧波-14");
                                TitleName.Add("S相電流諧波-15");
                                TitleName.Add("S相電流諧波-16");
                                TitleName.Add("S相電流諧波-17");
                                TitleName.Add("S相電流諧波-18");
                                TitleName.Add("S相電流諧波-19");
                                TitleName.Add("S相電流諧波-20");
                                TitleName.Add("S相電流諧波-21");
                                TitleName.Add("S相電流諧波-22");
                                TitleName.Add("S相電流諧波-23");
                                TitleName.Add("S相電流諧波-24");
                                TitleName.Add("S相電流諧波-25");
                                TitleName.Add("S相電流諧波-26");
                                TitleName.Add("S相電流諧波-27");
                                TitleName.Add("S相電流諧波-28");
                                TitleName.Add("S相電流諧波-29");
                                TitleName.Add("S相電流諧波-30");
                                TitleName.Add("S相電流諧波-31");
                            }
                            break;
                        case "T相電流諧波":
                            {
                                value += ",TI_HD" +
                                    ",TI_HD1" +
                                    ",TI_HD2" +
                                    ",TI_HD3" +
                                    ",TI_HD4" +
                                    ",TI_HD5" +
                                    ",TI_HD6" +
                                    ",TI_HD7" +
                                    ",TI_HD8" +
                                    ",TI_HD9" +
                                    ",TI_HD10" +
                                    ",TI_HD11" +
                                    ",TI_HD12" +
                                    ",TI_HD13" +
                                    ",TI_HD14" +
                                    ",TI_HD15" +
                                    ",TI_HD16" +
                                    ",TI_HD17" +
                                    ",TI_HD18" +
                                    ",TI_HD19" +
                                    ",TI_HD20" +
                                    ",TI_HD21" +
                                    ",TI_HD22" +
                                    ",TI_HD23" +
                                    ",TI_HD24" +
                                    ",TI_HD25" +
                                    ",TI_HD26" +
                                    ",TI_HD27" +
                                    ",TI_HD28" +
                                    ",TI_HD29" +
                                    ",TI_HD30" +
                                    ",TI_HD31";
                                TitleName.Add("T相電流諧波-Reserved");
                                TitleName.Add("T相電流諧波-1");
                                TitleName.Add("T相電流諧波-2");
                                TitleName.Add("T相電流諧波-3");
                                TitleName.Add("T相電流諧波-4");
                                TitleName.Add("T相電流諧波-5");
                                TitleName.Add("T相電流諧波-6");
                                TitleName.Add("T相電流諧波-7");
                                TitleName.Add("T相電流諧波-8");
                                TitleName.Add("T相電流諧波-9");
                                TitleName.Add("T相電流諧波-10");
                                TitleName.Add("T相電流諧波-11");
                                TitleName.Add("T相電流諧波-12");
                                TitleName.Add("T相電流諧波-13");
                                TitleName.Add("T相電流諧波-14");
                                TitleName.Add("T相電流諧波-15");
                                TitleName.Add("T相電流諧波-16");
                                TitleName.Add("T相電流諧波-17");
                                TitleName.Add("T相電流諧波-18");
                                TitleName.Add("T相電流諧波-19");
                                TitleName.Add("T相電流諧波-20");
                                TitleName.Add("T相電流諧波-21");
                                TitleName.Add("T相電流諧波-22");
                                TitleName.Add("T相電流諧波-23");
                                TitleName.Add("T相電流諧波-24");
                                TitleName.Add("T相電流諧波-25");
                                TitleName.Add("T相電流諧波-26");
                                TitleName.Add("T相電流諧波-27");
                                TitleName.Add("T相電流諧波-28");
                                TitleName.Add("T相電流諧波-29");
                                TitleName.Add("T相電流諧波-30");
                                TitleName.Add("T相電流諧波-31");
                            }
                            break;
                    }
                    string sql = $"USE [EwatchIntegrate_Log] SELECT  {value} FROM Electric_{deviceSetting.CardNumber}{deviceSetting.BordNumber} WHERE DeviceNumber = {deviceSetting.DeviceNumber} AND (ttime >= '{StartTime:yyyyMMdd000000}' AND ttime <= '{EndTime:yyyyMMdd235959}')";
                    var data = connection.ExecuteReader(sql);
                    DataTable dt = new DataTable();
                    dt.Load(data);
                    IWorkbook wb = new XSSFWorkbook();
                    ISheet ws;
                    ws = wb.CreateSheet(Kind);
                    #region 設計標頭
                    ws.CreateRow(0);    //第一行為欄位名稱
                    XSSFCellStyle tStyle = (XSSFCellStyle)wb.CreateCellStyle();
                    tStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
                    tStyle.FillPattern = FillPattern.SolidForeground;
                    tStyle.BorderBottom = BorderStyle.Thin;
                    tStyle.BorderLeft = BorderStyle.Thin;
                    tStyle.BorderRight = BorderStyle.Thin;
                    tStyle.BorderTop = BorderStyle.Thin;
                    int titleCount = 0;
                    foreach (string titlename in TitleName)
                    {
                        ws.GetRow(0).CreateCell(titleCount).SetCellValue(titlename);
                        ws.GetRow(0).GetCell(titleCount).CellStyle = tStyle;
                        titleCount++;
                    }
                    #endregion
                    #region 設定明細
                    XSSFCellStyle xStyle = (XSSFCellStyle)wb.CreateCellStyle();
                    xStyle.BorderBottom = BorderStyle.Thin;
                    xStyle.BorderLeft = BorderStyle.Thin;
                    xStyle.BorderRight = BorderStyle.Thin;
                    xStyle.BorderTop = BorderStyle.Thin;
                    int rowCount = 1;

                    foreach (DataRow s in dt.Rows)
                    {
                        ws.CreateRow(rowCount);
                        for (int i = 0; i < s.ItemArray.Length; i++)
                        {
                            ws.GetRow(rowCount).CreateCell(i).SetCellValue(s.ItemArray[i].ToString());
                            ws.GetRow(rowCount).GetCell(i).CellStyle = xStyle;
                        }
                        rowCount++;
                    }
                    #endregion
                    NpoiMemoryStream stream = new NpoiMemoryStream();
                    stream.AllowClose = false;
                    wb.Write(stream, false);
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.AllowClose = true;
                    return new FileStreamResult(stream, $"application/xlsx") { FileDownloadName = $"{deviceSetting.DeviceName}分析表_{DateTimeOffset.Now.ToUnixTimeSeconds()}.xlsx" };
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return null;
            }
        }

        #endregion
        #region 事件查詢
        /// <summary>
        /// 事件報表查詢
        /// </summary>
        /// <param name="StartTime">開始時間</param>
        /// <param name="EndTime">結束時間</param>
        /// <param name="deviceSetting">設備資訊</param>
        /// <param name="Kind">種類</param>
        /// <returns></returns>
        public object Search_Event(DateTime StartTime, DateTime EndTime, DeviceSetting deviceSetting, string Kind)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    List<string> TitleName = new List<string>();
                    List<object> _Logs = new List<object>();
                    string sql = "";
                    switch (Kind)
                    {
                        case "驟降驟升事件":
                            {
                                sql = $"USE [EwatchIntegrate_Log] SELECT  ttimen,Sag_Swell_Type,Cycle,Data,Phase,Sag_Swell_Begin,Sag_Swell_End FROM Electric_Sag_Swell_{deviceSetting.CardNumber}{deviceSetting.BordNumber} WHERE DeviceNumber = {deviceSetting.DeviceNumber} AND (ttime >= '{StartTime:yyyyMMdd000000}' AND ttime <= '{EndTime:yyyyMMdd235959}')";
                                TitleName.Add("時間");
                                TitleName.Add("驟降驟升類型");
                                TitleName.Add("電壓驟降驟升之持續 cycle 數");
                                TitleName.Add("電壓驟降驟升發生時百分比");
                                TitleName.Add("相電壓");
                                TitleName.Add("開始發生日期與時間");
                                TitleName.Add("結束日期與時間");
                            }
                            break;
                        case "需量/告警紀錄":
                            {
                                sql = $"USE [EwatchIntegrate_Log] SELECT  ttimen,AlarmItem,AlarmData,AlarmTime FROM ElectricAlarm_{deviceSetting.CardNumber}{deviceSetting.BordNumber} WHERE DeviceNumber = {deviceSetting.DeviceNumber} AND (ttime >= '{StartTime:yyyyMMdd000000}' AND ttime <= '{EndTime:yyyyMMdd235959}')";
                                TitleName.Add("時間");
                                TitleName.Add("警報項目");
                                TitleName.Add("警報發生或解除時之百分比");
                                TitleName.Add("發生日期與時間");
                            }
                            break;
                    }
                    _Logs = connection.Query<object>(sql).ToList();
                    return new
                    {
                        TitleName = TitleName,
                        data = _Logs
                    };
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return new();
            }
        }
        /// <summary>
        /// 下載事件報表
        /// </summary>
        /// <param name="StartTime">開始時間</param>
        /// <param name="EndTime">結束時間</param>
        /// <param name="deviceSetting">設備資訊</param>
        /// <param name="Kind">種類</param>
        /// <returns></returns>
        public FileStreamResult? DownLoad_Device_Event_Report(DateTime StartTime, DateTime EndTime, DeviceSetting deviceSetting, string Kind)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    List<string> TitleName = new List<string>();
                    string sql = "";
                    switch (Kind)
                    {
                        case "驟降驟升事件":
                            {
                                sql = $"USE [EwatchIntegrate_Log] SELECT  ttimen,Sag_Swell_Type,Cycle,Data,Phase,Sag_Swell_Begin,Sag_Swell_End FROM Electric_Sag_Swell_{deviceSetting.CardNumber}{deviceSetting.BordNumber} WHERE DeviceNumber = {deviceSetting.DeviceNumber} AND (ttime >= '{StartTime:yyyyMMdd000000}' AND ttime <= '{EndTime:yyyyMMdd235959}')";
                                TitleName.Add("時間");
                                TitleName.Add("驟降驟升類型");
                                TitleName.Add("電壓驟降驟升之持續 cycle 數");
                                TitleName.Add("電壓驟降驟升發生時百分比");
                                TitleName.Add("相電壓");
                                TitleName.Add("開始發生日期與時間");
                                TitleName.Add("結束日期與時間");
                            }
                            break;
                        case "需量&告警紀錄":
                            {
                                sql = $"USE [EwatchIntegrate_Log] SELECT  ttimen,AlarmItem,AlarmData,AlarmTime FROM ElectricAlarm_{deviceSetting.CardNumber}{deviceSetting.BordNumber} WHERE DeviceNumber = {deviceSetting.DeviceNumber} AND (ttime >= '{StartTime:yyyyMMdd000000}' AND ttime <= '{EndTime:yyyyMMdd235959}')";
                                TitleName.Add("時間");
                                TitleName.Add("警報項目");
                                TitleName.Add("警報發生或解除時之百分比");
                                TitleName.Add("發生日期與時間");
                            }
                            break;
                    }
                    var data = connection.ExecuteReader(sql);
                    DataTable dt = new DataTable();
                    dt.Load(data);
                    IWorkbook wb = new XSSFWorkbook();
                    ISheet ws;
                    ws = wb.CreateSheet(Kind);
                    #region 設計標頭
                    ws.CreateRow(0);    //第一行為欄位名稱
                    XSSFCellStyle tStyle = (XSSFCellStyle)wb.CreateCellStyle();
                    tStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
                    tStyle.FillPattern = FillPattern.SolidForeground;
                    tStyle.BorderBottom = BorderStyle.Thin;
                    tStyle.BorderLeft = BorderStyle.Thin;
                    tStyle.BorderRight = BorderStyle.Thin;
                    tStyle.BorderTop = BorderStyle.Thin;
                    int titleCount = 0;
                    foreach (string titlename in TitleName)
                    {
                        ws.GetRow(0).CreateCell(titleCount).SetCellValue(titlename);
                        ws.GetRow(0).GetCell(titleCount).CellStyle = tStyle;
                        titleCount++;
                    }
                    #endregion
                    #region 設定明細
                    XSSFCellStyle xStyle = (XSSFCellStyle)wb.CreateCellStyle();
                    xStyle.BorderBottom = BorderStyle.Thin;
                    xStyle.BorderLeft = BorderStyle.Thin;
                    xStyle.BorderRight = BorderStyle.Thin;
                    xStyle.BorderTop = BorderStyle.Thin;
                    int rowCount = 1;

                    foreach (DataRow s in dt.Rows)
                    {
                        ws.CreateRow(rowCount);
                        for (int i = 0; i < s.ItemArray.Length; i++)
                        {
                            ws.GetRow(rowCount).CreateCell(i).SetCellValue(s.ItemArray[i].ToString());
                            ws.GetRow(rowCount).GetCell(i).CellStyle = xStyle;
                        }
                        rowCount++;
                    }
                    #endregion
                    NpoiMemoryStream stream = new NpoiMemoryStream();
                    stream.AllowClose = false;
                    wb.Write(stream, false);
                    stream.Seek(0, SeekOrigin.Begin);
                    stream.AllowClose = true;
                    return new FileStreamResult(stream, $"application/xlsx") { FileDownloadName = $"{deviceSetting.DeviceName}事件表_{DateTimeOffset.Now.ToUnixTimeSeconds()}.xlsx" };
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return null;
            }
        }
        #endregion
        #region 新增資訊
        /// <summary>
        /// 新增電表歷史資訊
        /// </summary>
        /// <param name="setting"></param>
        /// <param name="CardNumber"></param>
        /// <param name="BordNumber"></param>
        /// <returns></returns>
        public ResponseTypeEnum ElectricInformation_Add(ElectricInformation_Log setting, string CardNumber, string BordNumber)
        {
            try
            {
                if (CardBord_Search(CardNumber, BordNumber))
                {
                    Create_DataBase(setting.ttimen);
                    Create_ElectricTable(setting.ttimen, CardNumber, BordNumber);
                    using (IDbConnection connection = new SqlConnection(SqlDB))
                    {
                        //string sql = $"USE [EwatchIntegrate_Log_{setting.ttimen:yyyyMM}] IF NOT EXISTS (SELECT * FROM  Electric_{CardNumber}{BordNumber} WHERE ttime = @ttime AND DeviceNumber = @DeviceNumber)" +
                        string sql = $"USE [EwatchIntegrate_Log] IF NOT EXISTS (SELECT * FROM  Electric_{CardNumber}{BordNumber} WHERE ttime = @ttime AND DeviceNumber = @DeviceNumber)" +
                            $"INSERT INTO Electric_{CardNumber}{BordNumber} (" +
                            $"ttime" +
                            $",ttimen" +
                            $",DeviceNumber" +
                            $",Loop" +
                            $",Va" +
                            $",Vb" +
                            $",Vc" +
                            $",Vnavg" +
                            $",Vab" +
                            $",Vbc" +
                            $",Vca" +
                            $",Vlavg" +
                            $",Ia" +
                            $",Ib" +
                            $",Ic" +
                            $",Iavg" +
                            $",Freq" +
                            $",PFa" +
                            $",PFb" +
                            $",PFc" +
                            $",PFavg" +
                            $",Kwa" +
                            $",Kwb" +
                            $",Kwc" +
                            $",Kwtot" +
                            $",KVARa" +
                            $",KVARb" +
                            $",KVARc" +
                            $",KVARtot" +
                            $",KVAa" +
                            $",KVAb" +
                            $",KVAc" +
                            $",KVAtot" +
                            $",PhaseAngleVa" +
                            $",PhaseAngleVb" +
                            $",PhaseAngleVc" +
                            $",PhaseAngleIa" +
                            $",PhaseAngleIb" +
                            $",PhaseAngleIc" +
                            $",RV_HD" +
                            $",RV_HD1" +
                            $",RV_HD2" +
                            $",RV_HD3" +
                            $",RV_HD4" +
                            $",RV_HD5" +
                            $",RV_HD6" +
                            $",RV_HD7" +
                            $",RV_HD8" +
                            $",RV_HD9" +
                            $",RV_HD10" +
                            $",RV_HD11" +
                            $",RV_HD12" +
                            $",RV_HD13" +
                            $",RV_HD14" +
                            $",RV_HD15" +
                            $",RV_HD16" +
                            $",RV_HD17" +
                            $",RV_HD18" +
                            $",RV_HD19" +
                            $",RV_HD20" +
                            $",RV_HD21" +
                            $",RV_HD22" +
                            $",RV_HD23" +
                            $",RV_HD24" +
                            $",RV_HD25" +
                            $",RV_HD26" +
                            $",RV_HD27" +
                            $",RV_HD28" +
                            $",RV_HD29" +
                            $",RV_HD30" +
                            $",RV_HD31" +
                            $",SV_HD" +
                            $",SV_HD1" +
                            $",SV_HD2" +
                            $",SV_HD3" +
                            $",SV_HD4" +
                            $",SV_HD5" +
                            $",SV_HD6" +
                            $",SV_HD7" +
                            $",SV_HD8" +
                            $",SV_HD9" +
                            $",SV_HD10" +
                            $",SV_HD11" +
                            $",SV_HD12" +
                            $",SV_HD13" +
                            $",SV_HD14" +
                            $",SV_HD15" +
                            $",SV_HD16" +
                            $",SV_HD17" +
                            $",SV_HD18" +
                            $",SV_HD19" +
                            $",SV_HD20" +
                            $",SV_HD21" +
                            $",SV_HD22" +
                            $",SV_HD23" +
                            $",SV_HD24" +
                            $",SV_HD25" +
                            $",SV_HD26" +
                            $",SV_HD27" +
                            $",SV_HD28" +
                            $",SV_HD29" +
                            $",SV_HD30" +
                            $",SV_HD31" +
                            $",TV_HD" +
                            $",TV_HD1" +
                            $",TV_HD2" +
                            $",TV_HD3" +
                            $",TV_HD4" +
                            $",TV_HD5" +
                            $",TV_HD6" +
                            $",TV_HD7" +
                            $",TV_HD8" +
                            $",TV_HD9" +
                            $",TV_HD10" +
                            $",TV_HD11" +
                            $",TV_HD12" +
                            $",TV_HD13" +
                            $",TV_HD14" +
                            $",TV_HD15" +
                            $",TV_HD16" +
                            $",TV_HD17" +
                            $",TV_HD18" +
                            $",TV_HD19" +
                            $",TV_HD20" +
                            $",TV_HD21" +
                            $",TV_HD22" +
                            $",TV_HD23" +
                            $",TV_HD24" +
                            $",TV_HD25" +
                            $",TV_HD26" +
                            $",TV_HD27" +
                            $",TV_HD28" +
                            $",TV_HD29" +
                            $",TV_HD30" +
                            $",TV_HD31" +
                            $",RI_HD" +
                            $",RI_HD1" +
                            $",RI_HD2" +
                            $",RI_HD3" +
                            $",RI_HD4" +
                            $",RI_HD5" +
                            $",RI_HD6" +
                            $",RI_HD7" +
                            $",RI_HD8" +
                            $",RI_HD9" +
                            $",RI_HD10" +
                            $",RI_HD11" +
                            $",RI_HD12" +
                            $",RI_HD13" +
                            $",RI_HD14" +
                            $",RI_HD15" +
                            $",RI_HD16" +
                            $",RI_HD17" +
                            $",RI_HD18" +
                            $",RI_HD19" +
                            $",RI_HD20" +
                            $",RI_HD21" +
                            $",RI_HD22" +
                            $",RI_HD23" +
                            $",RI_HD24" +
                            $",RI_HD25" +
                            $",RI_HD26" +
                            $",RI_HD27" +
                            $",RI_HD28" +
                            $",RI_HD29" +
                            $",RI_HD30" +
                            $",RI_HD31" +
                            $",SI_HD" +
                            $",SI_HD1" +
                            $",SI_HD2" +
                            $",SI_HD3" +
                            $",SI_HD4" +
                            $",SI_HD5" +
                            $",SI_HD6" +
                            $",SI_HD7" +
                            $",SI_HD8" +
                            $",SI_HD9" +
                            $",SI_HD10" +
                            $",SI_HD11" +
                            $",SI_HD12" +
                            $",SI_HD13" +
                            $",SI_HD14" +
                            $",SI_HD15" +
                            $",SI_HD16" +
                            $",SI_HD17" +
                            $",SI_HD18" +
                            $",SI_HD19" +
                            $",SI_HD20" +
                            $",SI_HD21" +
                            $",SI_HD22" +
                            $",SI_HD23" +
                            $",SI_HD24" +
                            $",SI_HD25" +
                            $",SI_HD26" +
                            $",SI_HD27" +
                            $",SI_HD28" +
                            $",SI_HD29" +
                            $",SI_HD30" +
                            $",SI_HD31" +
                            $",TI_HD" +
                            $",TI_HD1" +
                            $",TI_HD2" +
                            $",TI_HD3" +
                            $",TI_HD4" +
                            $",TI_HD5" +
                            $",TI_HD6" +
                            $",TI_HD7" +
                            $",TI_HD8" +
                            $",TI_HD9" +
                            $",TI_HD10" +
                            $",TI_HD11" +
                            $",TI_HD12" +
                            $",TI_HD13" +
                            $",TI_HD14" +
                            $",TI_HD15" +
                            $",TI_HD16" +
                            $",TI_HD17" +
                            $",TI_HD18" +
                            $",TI_HD19" +
                            $",TI_HD20" +
                            $",TI_HD21" +
                            $",TI_HD22" +
                            $",TI_HD23" +
                            $",TI_HD24" +
                            $",TI_HD25" +
                            $",TI_HD26" +
                            $",TI_HD27" +
                            $",TI_HD28" +
                            $",TI_HD29" +
                            $",TI_HD30" +
                            $",TI_HD31" +
                            $",KWH" +
                            $") VALUES (" +
                            $"@ttime" +
                            $",@ttimen" +
                            $",@DeviceNumber" +
                            $",@Loop" +
                            $",@Va" +
                            $",@Vb" +
                            $",@Vc" +
                            $",@Vnavg" +
                            $",@Vab" +
                            $",@Vbc" +
                            $",@Vca" +
                            $",@Vlavg" +
                            $",@Ia" +
                            $",@Ib" +
                            $",@Ic" +
                            $",@Iavg" +
                            $",@Freq" +
                            $",@KVARa" +
                            $",@KVARb" +
                            $",@KVARc" +
                            $",@KVARtot" +
                            $",@KVAa" +
                            $",@KVAb" +
                            $",@KVAc" +
                            $",@KVAtot" +
                            $",@Kwa" +
                            $",@Kwb" +
                            $",@Kwc" +
                            $",@Kwtot" +
                            $",@PFa" +
                            $",@PFb" +
                            $",@PFc" +
                            $",@PFavg" +
                            $",@PhaseAngleVa" +
                            $",@PhaseAngleVb" +
                            $",@PhaseAngleVc" +
                            $",@PhaseAngleIa" +
                            $",@PhaseAngleIb" +
                            $",@PhaseAngleIc" +
                            $",@RV_HD" +
                            $",@RV_HD1" +
                            $",@RV_HD2" +
                            $",@RV_HD3" +
                            $",@RV_HD4" +
                            $",@RV_HD5" +
                            $",@RV_HD6" +
                            $",@RV_HD7" +
                            $",@RV_HD8" +
                            $",@RV_HD9" +
                            $",@RV_HD10" +
                            $",@RV_HD11" +
                            $",@RV_HD12" +
                            $",@RV_HD13" +
                            $",@RV_HD14" +
                            $",@RV_HD15" +
                            $",@RV_HD16" +
                            $",@RV_HD17" +
                            $",@RV_HD18" +
                            $",@RV_HD19" +
                            $",@RV_HD20" +
                            $",@RV_HD21" +
                            $",@RV_HD22" +
                            $",@RV_HD23" +
                            $",@RV_HD24" +
                            $",@RV_HD25" +
                            $",@RV_HD26" +
                            $",@RV_HD27" +
                            $",@RV_HD28" +
                            $",@RV_HD29" +
                            $",@RV_HD30" +
                            $",@RV_HD31" +
                            $",@SV_HD" +
                            $",@SV_HD1" +
                            $",@SV_HD2" +
                            $",@SV_HD3" +
                            $",@SV_HD4" +
                            $",@SV_HD5" +
                            $",@SV_HD6" +
                            $",@SV_HD7" +
                            $",@SV_HD8" +
                            $",@SV_HD9" +
                            $",@SV_HD10" +
                            $",@SV_HD11" +
                            $",@SV_HD12" +
                            $",@SV_HD13" +
                            $",@SV_HD14" +
                            $",@SV_HD15" +
                            $",@SV_HD16" +
                            $",@SV_HD17" +
                            $",@SV_HD18" +
                            $",@SV_HD19" +
                            $",@SV_HD20" +
                            $",@SV_HD21" +
                            $",@SV_HD22" +
                            $",@SV_HD23" +
                            $",@SV_HD24" +
                            $",@SV_HD25" +
                            $",@SV_HD26" +
                            $",@SV_HD27" +
                            $",@SV_HD28" +
                            $",@SV_HD29" +
                            $",@SV_HD30" +
                            $",@SV_HD31" +
                            $",@TV_HD" +
                            $",@TV_HD1" +
                            $",@TV_HD2" +
                            $",@TV_HD3" +
                            $",@TV_HD4" +
                            $",@TV_HD5" +
                            $",@TV_HD6" +
                            $",@TV_HD7" +
                            $",@TV_HD8" +
                            $",@TV_HD9" +
                            $",@TV_HD10" +
                            $",@TV_HD11" +
                            $",@TV_HD12" +
                            $",@TV_HD13" +
                            $",@TV_HD14" +
                            $",@TV_HD15" +
                            $",@TV_HD16" +
                            $",@TV_HD17" +
                            $",@TV_HD18" +
                            $",@TV_HD19" +
                            $",@TV_HD20" +
                            $",@TV_HD21" +
                            $",@TV_HD22" +
                            $",@TV_HD23" +
                            $",@TV_HD24" +
                            $",@TV_HD25" +
                            $",@TV_HD26" +
                            $",@TV_HD27" +
                            $",@TV_HD28" +
                            $",@TV_HD29" +
                            $",@TV_HD30" +
                            $",@TV_HD31" +
                            $",@RI_HD" +
                            $",@RI_HD1" +
                            $",@RI_HD2" +
                            $",@RI_HD3" +
                            $",@RI_HD4" +
                            $",@RI_HD5" +
                            $",@RI_HD6" +
                            $",@RI_HD7" +
                            $",@RI_HD8" +
                            $",@RI_HD9" +
                            $",@RI_HD10" +
                            $",@RI_HD11" +
                            $",@RI_HD12" +
                            $",@RI_HD13" +
                            $",@RI_HD14" +
                            $",@RI_HD15" +
                            $",@RI_HD16" +
                            $",@RI_HD17" +
                            $",@RI_HD18" +
                            $",@RI_HD19" +
                            $",@RI_HD20" +
                            $",@RI_HD21" +
                            $",@RI_HD22" +
                            $",@RI_HD23" +
                            $",@RI_HD24" +
                            $",@RI_HD25" +
                            $",@RI_HD26" +
                            $",@RI_HD27" +
                            $",@RI_HD28" +
                            $",@RI_HD29" +
                            $",@RI_HD30" +
                            $",@RI_HD31" +
                            $",@SI_HD" +
                            $",@SI_HD1" +
                            $",@SI_HD2" +
                            $",@SI_HD3" +
                            $",@SI_HD4" +
                            $",@SI_HD5" +
                            $",@SI_HD6" +
                            $",@SI_HD7" +
                            $",@SI_HD8" +
                            $",@SI_HD9" +
                            $",@SI_HD10" +
                            $",@SI_HD11" +
                            $",@SI_HD12" +
                            $",@SI_HD13" +
                            $",@SI_HD14" +
                            $",@SI_HD15" +
                            $",@SI_HD16" +
                            $",@SI_HD17" +
                            $",@SI_HD18" +
                            $",@SI_HD19" +
                            $",@SI_HD20" +
                            $",@SI_HD21" +
                            $",@SI_HD22" +
                            $",@SI_HD23" +
                            $",@SI_HD24" +
                            $",@SI_HD25" +
                            $",@SI_HD26" +
                            $",@SI_HD27" +
                            $",@SI_HD28" +
                            $",@SI_HD29" +
                            $",@SI_HD30" +
                            $",@SI_HD31" +
                            $",@TI_HD" +
                            $",@TI_HD1" +
                            $",@TI_HD2" +
                            $",@TI_HD3" +
                            $",@TI_HD4" +
                            $",@TI_HD5" +
                            $",@TI_HD6" +
                            $",@TI_HD7" +
                            $",@TI_HD8" +
                            $",@TI_HD9" +
                            $",@TI_HD10" +
                            $",@TI_HD11" +
                            $",@TI_HD12" +
                            $",@TI_HD13" +
                            $",@TI_HD14" +
                            $",@TI_HD15" +
                            $",@TI_HD16" +
                            $",@TI_HD17" +
                            $",@TI_HD18" +
                            $",@TI_HD19" +
                            $",@TI_HD20" +
                            $",@TI_HD21" +
                            $",@TI_HD22" +
                            $",@TI_HD23" +
                            $",@TI_HD24" +
                            $",@TI_HD25" +
                            $",@TI_HD26" +
                            $",@TI_HD27" +
                            $",@TI_HD28" +
                            $",@TI_HD29" +
                            $",@TI_HD30" +
                            $",@TI_HD31" +
                            $",@KWH)";
                        var Index = connection.Execute(sql, setting);
                        if (Index > 0)
                        {
                            return ResponseTypeEnum.Finish;
                        }
                        else
                        {
                            return ResponseTypeEnum.Repeat;
                        }
                    }
                }
                else
                {
                    return ResponseTypeEnum.NoneCardBord;
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return ResponseTypeEnum.Error;
            }
        }
        /// <summary>
        /// 新增或更新電表即時資訊(Web)
        /// </summary>
        /// <param name="setting"></param>
        /// <param name="CardNumber"></param>
        /// <param name="BordNumber"></param>
        /// <returns></returns>
        public ResponseTypeEnum ElectricWeb_AddUpdate(ElectricInformation_Log setting, string CardNumber, string BordNumber)
        {
            try
            {
                if (CardBord_Search(CardNumber, BordNumber))
                {
                    using (IDbConnection connection = new SqlConnection(SqlDB))
                    {
                        string sql = $" IF NOT EXISTS (SELECT * FROM  ElectricWeb WHERE CardNumber = '{CardNumber}' AND BordNumber = '{BordNumber}' AND DeviceNumber = @DeviceNumber)" +
                            $"INSERT INTO ElectricWeb (" +
                            $"ttime" +
                            $",ttimen" +
                            $",DeviceNumber" +
                            $",Loop" +
                            $",CardNumber" +
                            $",BordNumber" +
                            $",Va" +
                            $",Vb" +
                            $",Vc" +
                            $",Vnavg" +
                            $",Vab" +
                            $",Vbc" +
                            $",Vca" +
                            $",Vlavg" +
                            $",Ia" +
                            $",Ib" +
                            $",Ic" +
                            $",Iavg" +
                            $",Freq" +
                            $",PFa" +
                            $",PFb" +
                            $",PFc" +
                            $",PFavg" +
                            $",Kwa" +
                            $",Kwb" +
                            $",Kwc" +
                            $",Kwtot" +
                            $",KVARa" +
                            $",KVARb" +
                            $",KVARc" +
                            $",KVARtot" +
                            $",KVAa" +
                            $",KVAb" +
                            $",KVAc" +
                            $",KVAtot" +
                            $",PhaseAngleVa" +
                            $",PhaseAngleVb" +
                            $",PhaseAngleVc" +
                            $",PhaseAngleIa" +
                            $",PhaseAngleIb" +
                            $",PhaseAngleIc" +
                            $",RV_HD" +
                            $",RV_HD1" +
                            $",RV_HD2" +
                            $",RV_HD3" +
                            $",RV_HD4" +
                            $",RV_HD5" +
                            $",RV_HD6" +
                            $",RV_HD7" +
                            $",RV_HD8" +
                            $",RV_HD9" +
                            $",RV_HD10" +
                            $",RV_HD11" +
                            $",RV_HD12" +
                            $",RV_HD13" +
                            $",RV_HD14" +
                            $",RV_HD15" +
                            $",RV_HD16" +
                            $",RV_HD17" +
                            $",RV_HD18" +
                            $",RV_HD19" +
                            $",RV_HD20" +
                            $",RV_HD21" +
                            $",RV_HD22" +
                            $",RV_HD23" +
                            $",RV_HD24" +
                            $",RV_HD25" +
                            $",RV_HD26" +
                            $",RV_HD27" +
                            $",RV_HD28" +
                            $",RV_HD29" +
                            $",RV_HD30" +
                            $",RV_HD31" +
                            $",SV_HD" +
                            $",SV_HD1" +
                            $",SV_HD2" +
                            $",SV_HD3" +
                            $",SV_HD4" +
                            $",SV_HD5" +
                            $",SV_HD6" +
                            $",SV_HD7" +
                            $",SV_HD8" +
                            $",SV_HD9" +
                            $",SV_HD10" +
                            $",SV_HD11" +
                            $",SV_HD12" +
                            $",SV_HD13" +
                            $",SV_HD14" +
                            $",SV_HD15" +
                            $",SV_HD16" +
                            $",SV_HD17" +
                            $",SV_HD18" +
                            $",SV_HD19" +
                            $",SV_HD20" +
                            $",SV_HD21" +
                            $",SV_HD22" +
                            $",SV_HD23" +
                            $",SV_HD24" +
                            $",SV_HD25" +
                            $",SV_HD26" +
                            $",SV_HD27" +
                            $",SV_HD28" +
                            $",SV_HD29" +
                            $",SV_HD30" +
                            $",SV_HD31" +
                            $",TV_HD" +
                            $",TV_HD1" +
                            $",TV_HD2" +
                            $",TV_HD3" +
                            $",TV_HD4" +
                            $",TV_HD5" +
                            $",TV_HD6" +
                            $",TV_HD7" +
                            $",TV_HD8" +
                            $",TV_HD9" +
                            $",TV_HD10" +
                            $",TV_HD11" +
                            $",TV_HD12" +
                            $",TV_HD13" +
                            $",TV_HD14" +
                            $",TV_HD15" +
                            $",TV_HD16" +
                            $",TV_HD17" +
                            $",TV_HD18" +
                            $",TV_HD19" +
                            $",TV_HD20" +
                            $",TV_HD21" +
                            $",TV_HD22" +
                            $",TV_HD23" +
                            $",TV_HD24" +
                            $",TV_HD25" +
                            $",TV_HD26" +
                            $",TV_HD27" +
                            $",TV_HD28" +
                            $",TV_HD29" +
                            $",TV_HD30" +
                            $",TV_HD31" +
                            $",RI_HD" +
                            $",RI_HD1" +
                            $",RI_HD2" +
                            $",RI_HD3" +
                            $",RI_HD4" +
                            $",RI_HD5" +
                            $",RI_HD6" +
                            $",RI_HD7" +
                            $",RI_HD8" +
                            $",RI_HD9" +
                            $",RI_HD10" +
                            $",RI_HD11" +
                            $",RI_HD12" +
                            $",RI_HD13" +
                            $",RI_HD14" +
                            $",RI_HD15" +
                            $",RI_HD16" +
                            $",RI_HD17" +
                            $",RI_HD18" +
                            $",RI_HD19" +
                            $",RI_HD20" +
                            $",RI_HD21" +
                            $",RI_HD22" +
                            $",RI_HD23" +
                            $",RI_HD24" +
                            $",RI_HD25" +
                            $",RI_HD26" +
                            $",RI_HD27" +
                            $",RI_HD28" +
                            $",RI_HD29" +
                            $",RI_HD30" +
                            $",RI_HD31" +
                            $",SI_HD" +
                            $",SI_HD1" +
                            $",SI_HD2" +
                            $",SI_HD3" +
                            $",SI_HD4" +
                            $",SI_HD5" +
                            $",SI_HD6" +
                            $",SI_HD7" +
                            $",SI_HD8" +
                            $",SI_HD9" +
                            $",SI_HD10" +
                            $",SI_HD11" +
                            $",SI_HD12" +
                            $",SI_HD13" +
                            $",SI_HD14" +
                            $",SI_HD15" +
                            $",SI_HD16" +
                            $",SI_HD17" +
                            $",SI_HD18" +
                            $",SI_HD19" +
                            $",SI_HD20" +
                            $",SI_HD21" +
                            $",SI_HD22" +
                            $",SI_HD23" +
                            $",SI_HD24" +
                            $",SI_HD25" +
                            $",SI_HD26" +
                            $",SI_HD27" +
                            $",SI_HD28" +
                            $",SI_HD29" +
                            $",SI_HD30" +
                            $",SI_HD31" +
                            $",TI_HD" +
                            $",TI_HD1" +
                            $",TI_HD2" +
                            $",TI_HD3" +
                            $",TI_HD4" +
                            $",TI_HD5" +
                            $",TI_HD6" +
                            $",TI_HD7" +
                            $",TI_HD8" +
                            $",TI_HD9" +
                            $",TI_HD10" +
                            $",TI_HD11" +
                            $",TI_HD12" +
                            $",TI_HD13" +
                            $",TI_HD14" +
                            $",TI_HD15" +
                            $",TI_HD16" +
                            $",TI_HD17" +
                            $",TI_HD18" +
                            $",TI_HD19" +
                            $",TI_HD20" +
                            $",TI_HD21" +
                            $",TI_HD22" +
                            $",TI_HD23" +
                            $",TI_HD24" +
                            $",TI_HD25" +
                            $",TI_HD26" +
                            $",TI_HD27" +
                            $",TI_HD28" +
                            $",TI_HD29" +
                            $",TI_HD30" +
                            $",TI_HD31" +
                            $",KWH" +
                            $") VALUES (" +
                            $"@ttime" +
                            $",@ttimen" +
                            $",@DeviceNumber" +
                            $",@Loop" +
                            $",'{CardNumber}'" +
                            $",'{BordNumber}'" +
                            $",@Va" +
                            $",@Vb" +
                            $",@Vc" +
                            $",@Vnavg" +
                            $",@Vab" +
                            $",@Vbc" +
                            $",@Vca" +
                            $",@Vlavg" +
                            $",@Ia" +
                            $",@Ib" +
                            $",@Ic" +
                            $",@Iavg" +
                            $",@Freq" +
                            $",@KVARa" +
                            $",@KVARb" +
                            $",@KVARc" +
                            $",@KVARtot" +
                            $",@KVAa" +
                            $",@KVAb" +
                            $",@KVAc" +
                            $",@KVAtot" +
                            $",@Kwa" +
                            $",@Kwb" +
                            $",@Kwc" +
                            $",@Kwtot" +
                            $",@PFa" +
                            $",@PFb" +
                            $",@PFc" +
                            $",@PFavg" +
                            $",@PhaseAngleVa" +
                            $",@PhaseAngleVb" +
                            $",@PhaseAngleVc" +
                            $",@PhaseAngleIa" +
                            $",@PhaseAngleIb" +
                            $",@PhaseAngleIc" +
                            $",@RV_HD" +
                            $",@RV_HD1" +
                            $",@RV_HD2" +
                            $",@RV_HD3" +
                            $",@RV_HD4" +
                            $",@RV_HD5" +
                            $",@RV_HD6" +
                            $",@RV_HD7" +
                            $",@RV_HD8" +
                            $",@RV_HD9" +
                            $",@RV_HD10" +
                            $",@RV_HD11" +
                            $",@RV_HD12" +
                            $",@RV_HD13" +
                            $",@RV_HD14" +
                            $",@RV_HD15" +
                            $",@RV_HD16" +
                            $",@RV_HD17" +
                            $",@RV_HD18" +
                            $",@RV_HD19" +
                            $",@RV_HD20" +
                            $",@RV_HD21" +
                            $",@RV_HD22" +
                            $",@RV_HD23" +
                            $",@RV_HD24" +
                            $",@RV_HD25" +
                            $",@RV_HD26" +
                            $",@RV_HD27" +
                            $",@RV_HD28" +
                            $",@RV_HD29" +
                            $",@RV_HD30" +
                            $",@RV_HD31" +
                            $",@SV_HD" +
                            $",@SV_HD1" +
                            $",@SV_HD2" +
                            $",@SV_HD3" +
                            $",@SV_HD4" +
                            $",@SV_HD5" +
                            $",@SV_HD6" +
                            $",@SV_HD7" +
                            $",@SV_HD8" +
                            $",@SV_HD9" +
                            $",@SV_HD10" +
                            $",@SV_HD11" +
                            $",@SV_HD12" +
                            $",@SV_HD13" +
                            $",@SV_HD14" +
                            $",@SV_HD15" +
                            $",@SV_HD16" +
                            $",@SV_HD17" +
                            $",@SV_HD18" +
                            $",@SV_HD19" +
                            $",@SV_HD20" +
                            $",@SV_HD21" +
                            $",@SV_HD22" +
                            $",@SV_HD23" +
                            $",@SV_HD24" +
                            $",@SV_HD25" +
                            $",@SV_HD26" +
                            $",@SV_HD27" +
                            $",@SV_HD28" +
                            $",@SV_HD29" +
                            $",@SV_HD30" +
                            $",@SV_HD31" +
                            $",@TV_HD" +
                            $",@TV_HD1" +
                            $",@TV_HD2" +
                            $",@TV_HD3" +
                            $",@TV_HD4" +
                            $",@TV_HD5" +
                            $",@TV_HD6" +
                            $",@TV_HD7" +
                            $",@TV_HD8" +
                            $",@TV_HD9" +
                            $",@TV_HD10" +
                            $",@TV_HD11" +
                            $",@TV_HD12" +
                            $",@TV_HD13" +
                            $",@TV_HD14" +
                            $",@TV_HD15" +
                            $",@TV_HD16" +
                            $",@TV_HD17" +
                            $",@TV_HD18" +
                            $",@TV_HD19" +
                            $",@TV_HD20" +
                            $",@TV_HD21" +
                            $",@TV_HD22" +
                            $",@TV_HD23" +
                            $",@TV_HD24" +
                            $",@TV_HD25" +
                            $",@TV_HD26" +
                            $",@TV_HD27" +
                            $",@TV_HD28" +
                            $",@TV_HD29" +
                            $",@TV_HD30" +
                            $",@TV_HD31" +
                            $",@RI_HD" +
                            $",@RI_HD1" +
                            $",@RI_HD2" +
                            $",@RI_HD3" +
                            $",@RI_HD4" +
                            $",@RI_HD5" +
                            $",@RI_HD6" +
                            $",@RI_HD7" +
                            $",@RI_HD8" +
                            $",@RI_HD9" +
                            $",@RI_HD10" +
                            $",@RI_HD11" +
                            $",@RI_HD12" +
                            $",@RI_HD13" +
                            $",@RI_HD14" +
                            $",@RI_HD15" +
                            $",@RI_HD16" +
                            $",@RI_HD17" +
                            $",@RI_HD18" +
                            $",@RI_HD19" +
                            $",@RI_HD20" +
                            $",@RI_HD21" +
                            $",@RI_HD22" +
                            $",@RI_HD23" +
                            $",@RI_HD24" +
                            $",@RI_HD25" +
                            $",@RI_HD26" +
                            $",@RI_HD27" +
                            $",@RI_HD28" +
                            $",@RI_HD29" +
                            $",@RI_HD30" +
                            $",@RI_HD31" +
                            $",@SI_HD" +
                            $",@SI_HD1" +
                            $",@SI_HD2" +
                            $",@SI_HD3" +
                            $",@SI_HD4" +
                            $",@SI_HD5" +
                            $",@SI_HD6" +
                            $",@SI_HD7" +
                            $",@SI_HD8" +
                            $",@SI_HD9" +
                            $",@SI_HD10" +
                            $",@SI_HD11" +
                            $",@SI_HD12" +
                            $",@SI_HD13" +
                            $",@SI_HD14" +
                            $",@SI_HD15" +
                            $",@SI_HD16" +
                            $",@SI_HD17" +
                            $",@SI_HD18" +
                            $",@SI_HD19" +
                            $",@SI_HD20" +
                            $",@SI_HD21" +
                            $",@SI_HD22" +
                            $",@SI_HD23" +
                            $",@SI_HD24" +
                            $",@SI_HD25" +
                            $",@SI_HD26" +
                            $",@SI_HD27" +
                            $",@SI_HD28" +
                            $",@SI_HD29" +
                            $",@SI_HD30" +
                            $",@SI_HD31" +
                            $",@TI_HD" +
                            $",@TI_HD1" +
                            $",@TI_HD2" +
                            $",@TI_HD3" +
                            $",@TI_HD4" +
                            $",@TI_HD5" +
                            $",@TI_HD6" +
                            $",@TI_HD7" +
                            $",@TI_HD8" +
                            $",@TI_HD9" +
                            $",@TI_HD10" +
                            $",@TI_HD11" +
                            $",@TI_HD12" +
                            $",@TI_HD13" +
                            $",@TI_HD14" +
                            $",@TI_HD15" +
                            $",@TI_HD16" +
                            $",@TI_HD17" +
                            $",@TI_HD18" +
                            $",@TI_HD19" +
                            $",@TI_HD20" +
                            $",@TI_HD21" +
                            $",@TI_HD22" +
                            $",@TI_HD23" +
                            $",@TI_HD24" +
                            $",@TI_HD25" +
                            $",@TI_HD26" +
                            $",@TI_HD27" +
                            $",@TI_HD28" +
                            $",@TI_HD29" +
                            $",@TI_HD30" +
                            $",@TI_HD31" +
                            $",@KWH) " +
                            $"ELSE " +
                            $"UPDATE ElectricWeb SET " +
                            $"ttime = @ttime" +
                            $",ttimen = @ttimen" +
                            $",Va=@Va" +
                            $",Vb=@Vb" +
                            $",Vc=@Vc" +
                            $",Vnavg=@Vnavg" +
                            $",Vab=@Vab" +
                            $",Vbc=@Vbc" +
                            $",Vca=@Vca" +
                            $",Vlavg=@Vlavg" +
                            $",Ia=@Ia" +
                            $",Ib=@Ib" +
                            $",Ic=@Ic" +
                            $",Iavg=@Iavg" +
                            $",Freq=@Freq" +
                            $",KVARa=@KVARa" +
                            $",KVARb=@KVARb" +
                            $",KVARc=@KVARc" +
                            $",KVARtot=@KVARtot" +
                            $",KVAa=@KVAa" +
                            $",KVAb=@KVAb" +
                            $",KVAc=@KVAc" +
                            $",KVAtot=@KVAtot" +
                            $",PFa=@PFa" +
                            $",PFb=@PFb" +
                            $",PFc=@PFc" +
                            $",PFavg=@PFavg" +
                            $",Kwa=@Kwa" +
                            $",Kwb=@Kwb" +
                            $",Kwc=@Kwc" +
                            $",Kwtot=@Kwtot" +
                            $",PhaseAngleVa=@PhaseAngleVa" +
                            $",PhaseAngleVb=@PhaseAngleVb" +
                            $",PhaseAngleVc=@PhaseAngleVc" +
                            $",PhaseAngleIa=@PhaseAngleIa" +
                            $",PhaseAngleIb=@PhaseAngleIb" +
                            $",PhaseAngleIc=@PhaseAngleIc" +
                            $",RV_HD = @RV_HD" +
                            $",RV_HD1=@RV_HD1" +
                            $",RV_HD2=@RV_HD2" +
                            $",RV_HD3=@RV_HD3" +
                            $",RV_HD4=@RV_HD4" +
                            $",RV_HD5=@RV_HD5" +
                            $",RV_HD6=@RV_HD6" +
                            $",RV_HD7=@RV_HD7" +
                            $",RV_HD8=@RV_HD8" +
                            $",RV_HD9=@RV_HD9" +
                            $",RV_HD10=@RV_HD10" +
                            $",RV_HD11=@RV_HD11" +
                            $",RV_HD12=@RV_HD12" +
                            $",RV_HD13=@RV_HD13" +
                            $",RV_HD14=@RV_HD14" +
                            $",RV_HD15=@RV_HD15" +
                            $",RV_HD16=@RV_HD16" +
                            $",RV_HD17=@RV_HD17" +
                            $",RV_HD18=@RV_HD18" +
                            $",RV_HD19=@RV_HD19" +
                            $",RV_HD20=@RV_HD20" +
                            $",RV_HD21=@RV_HD21" +
                            $",RV_HD22=@RV_HD22" +
                            $",RV_HD23=@RV_HD23" +
                            $",RV_HD24=@RV_HD24" +
                            $",RV_HD25=@RV_HD25" +
                            $",RV_HD26=@RV_HD26" +
                            $",RV_HD27=@RV_HD27" +
                            $",RV_HD28=@RV_HD28" +
                            $",RV_HD29=@RV_HD29" +
                            $",RV_HD30=@RV_HD30" +
                            $",RV_HD31=@RV_HD31" +
                            $",SV_HD=@SV_HD" +
                            $",SV_HD1=@SV_HD1" +
                            $",SV_HD2=@SV_HD2" +
                            $",SV_HD3=@SV_HD3" +
                            $",SV_HD4=@SV_HD4" +
                            $",SV_HD5=@SV_HD5" +
                            $",SV_HD6=@SV_HD6" +
                            $",SV_HD7=@SV_HD7" +
                            $",SV_HD8=@SV_HD8" +
                            $",SV_HD9=@SV_HD9" +
                            $",SV_HD10=@SV_HD10" +
                            $",SV_HD11=@SV_HD11" +
                            $",SV_HD12=@SV_HD12" +
                            $",SV_HD13=@SV_HD13" +
                            $",SV_HD14=@SV_HD14" +
                            $",SV_HD15=@SV_HD15" +
                            $",SV_HD16=@SV_HD16" +
                            $",SV_HD17=@SV_HD17" +
                            $",SV_HD18=@SV_HD18" +
                            $",SV_HD19=@SV_HD19" +
                            $",SV_HD20=@SV_HD20" +
                            $",SV_HD21=@SV_HD21" +
                            $",SV_HD22=@SV_HD22" +
                            $",SV_HD23=@SV_HD23" +
                            $",SV_HD24=@SV_HD24" +
                            $",SV_HD25=@SV_HD25" +
                            $",SV_HD26=@SV_HD26" +
                            $",SV_HD27=@SV_HD27" +
                            $",SV_HD28=@SV_HD28" +
                            $",SV_HD29=@SV_HD29" +
                            $",SV_HD30=@SV_HD30" +
                            $",SV_HD31=@SV_HD31" +
                            $",TV_HD=@TV_HD" +
                            $",TV_HD1=@TV_HD1" +
                            $",TV_HD2=@TV_HD2" +
                            $",TV_HD3=@TV_HD3" +
                            $",TV_HD4=@TV_HD4" +
                            $",TV_HD5=@TV_HD5" +
                            $",TV_HD6=@TV_HD6" +
                            $",TV_HD7=@TV_HD7" +
                            $",TV_HD8=@TV_HD8" +
                            $",TV_HD9=@TV_HD9" +
                            $",TV_HD10=@TV_HD10" +
                            $",TV_HD11=@TV_HD11" +
                            $",TV_HD12=@TV_HD12" +
                            $",TV_HD13=@TV_HD13" +
                            $",TV_HD14=@TV_HD14" +
                            $",TV_HD15=@TV_HD15" +
                            $",TV_HD16=@TV_HD16" +
                            $",TV_HD17=@TV_HD17" +
                            $",TV_HD18=@TV_HD18" +
                            $",TV_HD19=@TV_HD19" +
                            $",TV_HD20=@TV_HD20" +
                            $",TV_HD21=@TV_HD21" +
                            $",TV_HD22=@TV_HD22" +
                            $",TV_HD23=@TV_HD23" +
                            $",TV_HD24=@TV_HD24" +
                            $",TV_HD25=@TV_HD25" +
                            $",TV_HD26=@TV_HD26" +
                            $",TV_HD27=@TV_HD27" +
                            $",TV_HD28=@TV_HD28" +
                            $",TV_HD29=@TV_HD29" +
                            $",TV_HD30=@TV_HD30" +
                            $",TV_HD31=@TV_HD31" +
                            $",RI_HD =@RI_HD" +
                            $",RI_HD1=@RI_HD1" +
                            $",RI_HD2=@RI_HD2" +
                            $",RI_HD3=@RI_HD3" +
                            $",RI_HD4=@RI_HD4" +
                            $",RI_HD5=@RI_HD5" +
                            $",RI_HD6=@RI_HD6" +
                            $",RI_HD7=@RI_HD7" +
                            $",RI_HD8=@RI_HD8" +
                            $",RI_HD9=@RI_HD9" +
                            $",RI_HD10=@RI_HD10" +
                            $",RI_HD11=@RI_HD11" +
                            $",RI_HD12=@RI_HD12" +
                            $",RI_HD13=@RI_HD13" +
                            $",RI_HD14=@RI_HD14" +
                            $",RI_HD15=@RI_HD15" +
                            $",RI_HD16=@RI_HD16" +
                            $",RI_HD17=@RI_HD17" +
                            $",RI_HD18=@RI_HD18" +
                            $",RI_HD19=@RI_HD19" +
                            $",RI_HD20=@RI_HD20" +
                            $",RI_HD21=@RI_HD21" +
                            $",RI_HD22=@RI_HD22" +
                            $",RI_HD23=@RI_HD23" +
                            $",RI_HD24=@RI_HD24" +
                            $",RI_HD25=@RI_HD25" +
                            $",RI_HD26=@RI_HD26" +
                            $",RI_HD27=@RI_HD27" +
                            $",RI_HD28=@RI_HD28" +
                            $",RI_HD29=@RI_HD29" +
                            $",RI_HD30=@RI_HD30" +
                            $",RI_HD31=@RI_HD31" +
                            $",SI_HD =@SI_HD" +
                            $",SI_HD1=@SI_HD1" +
                            $",SI_HD2=@SI_HD2" +
                            $",SI_HD3=@SI_HD3" +
                            $",SI_HD4=@SI_HD4" +
                            $",SI_HD5=@SI_HD5" +
                            $",SI_HD6=@SI_HD6" +
                            $",SI_HD7=@SI_HD7" +
                            $",SI_HD8=@SI_HD8" +
                            $",SI_HD9=@SI_HD9" +
                            $",SI_HD10=@SI_HD10" +
                            $",SI_HD11=@SI_HD11" +
                            $",SI_HD12=@SI_HD12" +
                            $",SI_HD13=@SI_HD13" +
                            $",SI_HD14=@SI_HD14" +
                            $",SI_HD15=@SI_HD15" +
                            $",SI_HD16=@SI_HD16" +
                            $",SI_HD17=@SI_HD17" +
                            $",SI_HD18=@SI_HD18" +
                            $",SI_HD19=@SI_HD19" +
                            $",SI_HD20=@SI_HD20" +
                            $",SI_HD21=@SI_HD21" +
                            $",SI_HD22=@SI_HD22" +
                            $",SI_HD23=@SI_HD23" +
                            $",SI_HD24=@SI_HD24" +
                            $",SI_HD25=@SI_HD25" +
                            $",SI_HD26=@SI_HD26" +
                            $",SI_HD27=@SI_HD27" +
                            $",SI_HD28=@SI_HD28" +
                            $",SI_HD29=@SI_HD29" +
                            $",SI_HD30=@SI_HD30" +
                            $",SI_HD31=@SI_HD31" +
                            $",TI_HD =@TI_HD" +
                            $",TI_HD1=@TI_HD1" +
                            $",TI_HD2=@TI_HD2" +
                            $",TI_HD3=@TI_HD3" +
                            $",TI_HD4=@TI_HD4" +
                            $",TI_HD5=@TI_HD5" +
                            $",TI_HD6=@TI_HD6" +
                            $",TI_HD7=@TI_HD7" +
                            $",TI_HD8=@TI_HD8" +
                            $",TI_HD9=@TI_HD9" +
                            $",TI_HD10=@TI_HD10" +
                            $",TI_HD11=@TI_HD11" +
                            $",TI_HD12=@TI_HD12" +
                            $",TI_HD13=@TI_HD13" +
                            $",TI_HD14=@TI_HD14" +
                            $",TI_HD15=@TI_HD15" +
                            $",TI_HD16=@TI_HD16" +
                            $",TI_HD17=@TI_HD17" +
                            $",TI_HD18=@TI_HD18" +
                            $",TI_HD19=@TI_HD19" +
                            $",TI_HD20=@TI_HD20" +
                            $",TI_HD21=@TI_HD21" +
                            $",TI_HD22=@TI_HD22" +
                            $",TI_HD23=@TI_HD23" +
                            $",TI_HD24=@TI_HD24" +
                            $",TI_HD25=@TI_HD25" +
                            $",TI_HD26=@TI_HD26" +
                            $",TI_HD27=@TI_HD27" +
                            $",TI_HD28=@TI_HD28" +
                            $",TI_HD29=@TI_HD29" +
                            $",TI_HD30=@TI_HD30" +
                            $",TI_HD31=@TI_HD31" +
                            $",KWH=@KWH" +
                            $" WHERE DeviceNumber = @DeviceNumber AND CardNumber = '{CardNumber}' AND BordNumber = '{BordNumber}'";
                        var Index = connection.Execute(sql, setting);
                        if (Index > 0)
                        {
                            return ResponseTypeEnum.Finish;
                        }
                        else
                        {
                            return ResponseTypeEnum.Repeat;
                        }
                    }
                }
                else
                {
                    return ResponseTypeEnum.NoneCardBord;
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return ResponseTypeEnum.Error;
            }
        }
        /// <summary>
        /// 新增電表警告歷史資訊
        /// </summary>
        /// <param name="setting"></param>
        /// <param name="CardNumber"></param>
        /// <param name="BordNumber"></param>
        /// <returns></returns>
        public ResponseTypeEnum ElectricAlarm_Add(ElectricAlarm_Log setting, string CardNumber, string BordNumber)
        {
            try
            {
                if (CardBord_Search(CardNumber, BordNumber))
                {
                    Create_DataBase(setting.ttimen);
                    //Create_ElectricTable(setting.ttimen, CardNumber, BordNumber);
                    //Create_ElectricSagTable(setting.ttimen, CardNumber, BordNumber);
                    //Create_ElectricSwellTable(setting.ttimen, CardNumber, BordNumber);
                    Create_ElectricAlarmTable(setting.ttimen, CardNumber, BordNumber);
                    using (IDbConnection connection = new SqlConnection(SqlDB))
                    {
                        string sql = $"USE [EwatchIntegrate_Log] IF NOT EXISTS (SELECT * FROM  ElectricAlarm_{CardNumber}{BordNumber} WHERE ttime = @ttime AND DeviceNumber = @DeviceNumber)" +
                            $"INSERT INTO ElectricAlarm_{CardNumber}{BordNumber} (" +
                            $"ttime" +
                            $",ttimen" +
                            $",DeviceNumber" +
                            $",AlarmItem" +
                            $",AlarmData" +
                            $",AlarmTime" +
                            $") VALUES (" +
                            $"@ttime" +
                            $",@ttimen" +
                            $",@DeviceNumber" +
                            $",@AlarmItem" +
                            $",@AlarmData" +
                            $",@AlarmTime)";
                        var Index = connection.Execute(sql, setting);
                        if (Index > 0)
                        {
                            return ResponseTypeEnum.Finish;
                        }
                        else
                        {
                            return ResponseTypeEnum.Repeat;
                        }
                    }
                }
                else
                {
                    return ResponseTypeEnum.NoneCardBord;
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return ResponseTypeEnum.Error;
            }
        }
        /// <summary>
        /// 新增電表驟降驟升歷史資訊
        /// </summary>
        /// <param name="setting"></param>
        /// <param name="CardNumber"></param>
        /// <param name="BordNumber"></param>
        /// <returns></returns>
        public ResponseTypeEnum Electric_Sag_Swell_Add(Electric_Sag_Swell_Log setting, string CardNumber, string BordNumber)
        {
            try
            {
                if (CardBord_Search(CardNumber, BordNumber))
                {
                    Create_DataBase(setting.ttimen);
                    //Create_ElectricTable(setting.ttimen, CardNumber, BordNumber);
                    Create_Electric_Sag_SwellTable(setting.ttimen, CardNumber, BordNumber);
                    //Create_ElectricSwellTable(setting.ttimen, CardNumber, BordNumber);
                    //Create_ElectricAlarmTable(setting.ttimen, CardNumber, BordNumber);
                    using (IDbConnection connection = new SqlConnection(SqlDB))
                    {
                        //string sql = $"USE [EwatchIntegrate_Log_{setting.ttimen:yyyyMM}] IF NOT EXISTS (SELECT * FROM  ElectricSag_{CardNumber}{BordNumber} WHERE ttime = @ttime AND DeviceNumber = @DeviceNumber)" +
                        string sql = $"USE [EwatchIntegrate_Log] IF NOT EXISTS (SELECT * FROM  Electric_Sag_Swell_{CardNumber}{BordNumber} WHERE ttime = @ttime AND DeviceNumber = @DeviceNumber)" +
                            $"INSERT INTO Electric_Sag_Swell_{CardNumber}{BordNumber} (" +
                            $"ttime" +
                            $",ttimen" +
                            $",DeviceNumber" +
                            $",Sag_Swell_Type" +
                            $",Cycle" +
                            $",Data" +
                            $",Phase" +
                            $",Sag_Swell_Begin" +
                            $",Sag_Swell_End" +
                            $") VALUES (" +
                            $"@ttime" +
                            $",@ttimen" +
                            $",@DeviceNumber" +
                            $",@Sag_Swell_Type" +
                            $",@Cycle" +
                            $",@Data" +
                            $",@Phase" +
                            $",@Sag_Swell_Begin" +
                            $",@Sag_Swell_End)";
                        var Index = connection.Execute(sql, setting);
                        if (Index > 0)
                        {
                            return ResponseTypeEnum.Finish;
                        }
                        else
                        {
                            return ResponseTypeEnum.Repeat;
                        }
                    }
                }
                else
                {
                    return ResponseTypeEnum.NoneCardBord;
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return ResponseTypeEnum.Error;
            }
        }
        /// <summary>
        /// 累積電量預存程序
        /// </summary>
        /// <param name="setting"></param>
        /// <param name="CardNumber"></param>
        /// <param name="BordNumber"></param>
        /// <returns></returns>
        public ResponseTypeEnum ElectricTotal_Add(ElectricInformation_Log setting, string CardNumber, string BordNumber)
        {
            try
            {
                if (CardBord_Search(CardNumber, BordNumber))
                {
                    using (IDbConnection connection = new SqlConnection(SqlDB))
                    {
                        string sql = "USE [EwatchIntegrateTotal] EXEC KwhTotalProcedure @nowTime,@cardNumber,@bordNumber,@deviceNumber,@nowKwh";
                        var Index = connection.Execute(sql, new { nowTime = setting.ttime!.Substring(0, 8), cardNumber = CardNumber, bordNumber = BordNumber, deviceNumber = setting.DeviceNumber, nowKwh = setting.KWH });
                        if (Index > 0)
                        {
                            return ResponseTypeEnum.Finish;
                        }
                        else
                        {
                            return ResponseTypeEnum.Error;
                        }
                    }
                }
                else
                {
                    return ResponseTypeEnum.NoneCardBord;
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return ResponseTypeEnum.Error;
            }
        }
        #endregion
        #region NpoiMethod
        public class NpoiMemoryStream : MemoryStream
        {
            public bool AllowClose { get; set; }
            public NpoiMemoryStream()
            {
                AllowClose = true;
            }

            public override void Close()
            {
                if (AllowClose)
                {
                    base.Close();
                }
            }
        }
        #endregion
    }
}
