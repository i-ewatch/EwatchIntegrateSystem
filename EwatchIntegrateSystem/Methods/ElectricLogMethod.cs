using Dapper;
using EwatchIntegrateSystem.Enums;
using EwatchIntegrateSystem.Modules;
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
                    string sql = $"SELECT COUNT(*) FROM  master.dbo.sysdatabases WHERE name='EwatchIntegrate_{dateTime:yyyyMM}'";
                    int databaseflag = connection.QuerySingle<int>(sql);
                    if (databaseflag == 0)
                    {
                        string databaseSql = $"CREATE DATABASE [EwatchIntegrate_{dateTime:yyyyMM}]";
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
                    string sql = $"USE [EwatchIntegrate_{dateTime:yyyyMM}] SELECT COUNT(*) FROM  sys.tables WHERE name='Electric_{CardNumber}{BordNumber}'";
                    int Tableflag = connection.QuerySingle<int>(sql);
                    if (Tableflag == 0)
                    {
                        string dataTableSql = $"USE [EwatchIntegrate_{dateTime:yyyyMM}] CREATE TABLE [Electric_{CardNumber}{BordNumber}] (" +
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
                            $",[PFa] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PFb] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PFc] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PFavg] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PhaseAngleVa] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PhaseAngleVb] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PhaseAngleVc] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PhaseAngleIa] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PhaseAngleIb] DECIMAL(18,2) NOT NULL DEFAULT 0" +
                            $",[PhaseAngleIc] DECIMAL(18,2) NOT NULL DEFAULT 0" +
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
        /// 建立電表驟降事件表
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
                    string sql = $"USE [EwatchIntegrate_{dateTime:yyyyMM}] SELECT COUNT(*) FROM  sys.tables WHERE name='ElectricSag_{CardNumber}{BordNumber}'";
                    int Tableflag = connection.QuerySingle<int>(sql);
                    if (Tableflag == 0)
                    {
                        string dataTableSql = $"USE [EwatchIntegrate_{dateTime:yyyyMM}] CREATE TABLE [ElectricSag_{CardNumber}{BordNumber}] (" +
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
        /// 建立電表驟升事件表
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
                    string sql = $"USE [EwatchIntegrate_{dateTime:yyyyMM}] SELECT COUNT(*) FROM  sys.tables WHERE name='ElectricSwell_{CardNumber}{BordNumber}'";
                    int Tableflag = connection.QuerySingle<int>(sql);
                    if (Tableflag == 0)
                    {
                        string dataTableSql = $"USE [EwatchIntegrate_{dateTime:yyyyMM}] CREATE TABLE [ElectricSwell_{CardNumber}{BordNumber}] (" +
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
                    string sql = $"USE [EwatchIntegrate_{dateTime:yyyyMM}] SELECT COUNT(*) FROM  sys.tables WHERE name='ElectricAlarm_{CardNumber}{BordNumber}'";
                    int Tableflag = connection.QuerySingle<int>(sql);
                    if (Tableflag == 0)
                    {
                        string dataTableSql = $"USE [EwatchIntegrate_{dateTime:yyyyMM}] CREATE TABLE [ElectricAlarm_{CardNumber}{BordNumber}] (" +
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
        #region 查詢資訊
        /// <summary>
        /// 查詢電表歷史資料資訊(當月)
        /// </summary>
        /// <param name="StartTime">開始時間</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        public List<ElectricInformation_Log> ElectricInformation_SearchMonth(string StartTime, string CardNumber, string BordNumber)
        {
            List<ElectricInformation_Log> setting = new List<ElectricInformation_Log>();
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = $"USE [EwatchIntegrate_{StartTime.Substring(0, 6)}] SELECT * FROM Electric_{CardNumber}{BordNumber}";
                    setting = connection.Query<ElectricInformation_Log>(sql).ToList();
                }
                return setting;
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return setting;
            }
        }
        /// <summary>
        /// 查詢電表累積量資料
        /// </summary>
        /// <param name="StartTime">開始時間</param>
        /// <param name="EndTime">結束時間</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <param name="DeviceNumber">設備編號</param>
        /// <returns></returns>
        public List<ElectricTotal_Log> ElectricTotal_Search(string StartTime, string EndTime, string CardNumber, string BordNumber, int DeviceNumber)
        {
            List<ElectricTotal_Log> setting = new List<ElectricTotal_Log>();
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = $"USE [EwatchIntegrateTotal] SELECT * FROM Electric_Log WHERE (ttime >= @StartTime AND ttime <= @EndTime) AND CardNumber = @CardNumber AND BordNumber = @BordNumber AND DeviceNumber = @DeviceNumber";
                    setting = connection.Query<ElectricTotal_Log>(sql, new { StartTime = StartTime, EndTime, CardNumber, BordNumber, DeviceNumber }).ToList();
                }
                return setting;
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return setting;
            }
        }
        /// <summary>
        /// 查詢電表驟降歷史資料資訊(當月)
        /// </summary>
        /// <param name="StartTime">開始時間</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        public List<ElectricSag_Log> ElectricSag_SearchMonth(string StartTime, string CardNumber, string BordNumber)
        {
            List<ElectricSag_Log> setting = new List<ElectricSag_Log>();
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = $"USE [EwatchIntegrate_{StartTime.Substring(0, 6)}] SELECT * FROM ElectricSag_{CardNumber}{BordNumber}";
                    setting = connection.Query<ElectricSag_Log>(sql).ToList();
                }
                return setting;
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return setting;
            }
        }
        /// <summary>
        /// 查詢電表驟降最新一筆
        /// </summary>
        /// <param name="StartTime">時間資訊</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        public ElectricSag_Log? ElectricSag_Search(string StartTime, string CardNumber, string BordNumber)
        {
            ElectricSag_Log? setting = null;
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = $"USE [EwatchIntegrate_{StartTime.Substring(0, 6)}] SELECT TOP 1 * FROM ElectricSag_{CardNumber}{BordNumber} ORDER BY ttimen DESC";
                    setting = connection.QuerySingle<ElectricSag_Log>(sql);
                }
                return setting;
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return setting;
            }
        }
        /// <summary>
        /// 查詢電表驟升歷史資料資訊(當月)
        /// </summary>
        /// <param name="StartTime">開始時間</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        public List<ElectricSwell_Log> ElectricSwell_SearchMonth(string StartTime, string CardNumber, string BordNumber)
        {
            List<ElectricSwell_Log> setting = new List<ElectricSwell_Log>();
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = $"USE [EwatchIntegrate_{StartTime.Substring(0, 6)}] SELECT * FROM ElectricSwell_{CardNumber}{BordNumber}";
                    setting = connection.Query<ElectricSwell_Log>(sql).ToList();
                }
                return setting;
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return setting;
            }
        }
        /// <summary>
        /// 查詢電表驟升最新一筆
        /// </summary>
        /// <param name="StartTime">時間資訊</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        public ElectricSwell_Log? ElectricSwell_Search(string StartTime, string CardNumber, string BordNumber)
        {
            ElectricSwell_Log? setting = null;
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = $"USE [EwatchIntegrate_{StartTime.Substring(0, 6)}] SELECT TOP 1 * FROM ElectricSwell_{CardNumber}{BordNumber} ORDER BY ttimen DESC";
                    setting = connection.QuerySingle<ElectricSwell_Log>(sql);
                }
                return setting;
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return setting;
            }
        }
        /// <summary>
        /// 查詢電表告警歷史資料資訊(當月)
        /// </summary>
        /// <param name="StartTime">開始時間</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        public List<ElectricAlarm_Log> ElectricAlarm_SearchMonth(string StartTime, string CardNumber, string BordNumber)
        {
            List<ElectricAlarm_Log> setting = new List<ElectricAlarm_Log>();
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = $"USE [EwatchIntegrate_{StartTime.Substring(0, 6)}] SELECT * FROM ElectricAlarm_{CardNumber}{BordNumber}";
                    setting = connection.Query<ElectricAlarm_Log>(sql).ToList();
                }
                return setting;
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return setting;
            }
        }
        /// <summary>
        /// 查詢電表告警最新一筆
        /// </summary>
        /// <param name="StartTime">時間資訊</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        public ElectricAlarm_Log? ElectricAlarm_Search(string StartTime, string CardNumber, string BordNumber)
        {
            ElectricAlarm_Log? setting = null;
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = $"USE [EwatchIntegrate_{StartTime.Substring(0, 6)}] SELECT TOP 1 * FROM ElectricAlarm_{CardNumber}{BordNumber} ORDER BY ttimen DESC";
                    setting = connection.QuerySingle<ElectricAlarm_Log>(sql);
                }
                return setting;
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return setting;
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
                    //Create_ElectricSagTable(setting.ttimen, CardNumber, BordNumber);
                    //Create_ElectricSwellTable(setting.ttimen, CardNumber, BordNumber);
                    //Create_ElectricAlarmTable(setting.ttimen, CardNumber, BordNumber);
                    using (IDbConnection connection = new SqlConnection(SqlDB))
                    {
                        string sql = $"USE [EwatchIntegrate_{setting.ttimen:yyyyMM}] IF NOT EXISTS (SELECT * FROM  Electric_{CardNumber}{BordNumber} WHERE ttime = @ttime AND DeviceNumber = @DeviceNumber)" +
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
                            $",PFa" +
                            $",PFb" +
                            $",PFc" +
                            $",PFavg" +
                            $",PhaseAngleVa" +
                            $",PhaseAngleVb" +
                            $",PhaseAngleVc" +
                            $",PhaseAngleIa" +
                            $",PhaseAngleIb" +
                            $",PhaseAngleIc" +
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
                            $",@Kwa" +
                            $",@Kwb" +
                            $",@Kwc" +
                            $",@Kwtot" +
                            $",@KVARa" +
                            $",@KVARb" +
                            $",@KVARc" +
                            $",@KVARtot" +
                            $",@KVAa" +
                            $",@KVAb" +
                            $",@KVAc" +
                            $",@KVAtot" +
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
                            $",PFa" +
                            $",PFb" +
                            $",PFc" +
                            $",PFavg" +
                            $",PhaseAngleVa" +
                            $",PhaseAngleVb" +
                            $",PhaseAngleVc" +
                            $",PhaseAngleIa" +
                            $",PhaseAngleIb" +
                            $",PhaseAngleIc" +
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
                            $",@Kwa" +
                            $",@Kwb" +
                            $",@Kwc" +
                            $",@Kwtot" +
                            $",@KVARa" +
                            $",@KVARb" +
                            $",@KVARc" +
                            $",@KVARtot" +
                            $",@KVAa" +
                            $",@KVAb" +
                            $",@KVAc" +
                            $",@KVAtot" +
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
                            $",Kwa=@Kwa" +
                            $",Kwb=@Kwb" +
                            $",Kwc=@Kwc" +
                            $",Kwtot=@Kwtot" +
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
                            $",PhaseAngleVa=@PhaseAngleVa" +
                            $",PhaseAngleVb=@PhaseAngleVb" +
                            $",PhaseAngleVc=@PhaseAngleVc" +
                            $",PhaseAngleIa=@PhaseAngleIa" +
                            $",PhaseAngleIb=@PhaseAngleIb" +
                            $",PhaseAngleIc=@PhaseAngleIc" +
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
        /// 新增電表驟降歷史資訊
        /// </summary>
        /// <param name="setting"></param>
        /// <param name="CardNumber"></param>
        /// <param name="BordNumber"></param>
        /// <returns></returns>
        public ResponseTypeEnum ElectricSag_Add(ElectricSag_Log setting, string CardNumber, string BordNumber)
        {
            try
            {
                if (CardBord_Search(CardNumber, BordNumber))
                {
                    Create_DataBase(setting.ttimen);
                    //Create_ElectricTable(setting.ttimen, CardNumber, BordNumber);
                    Create_ElectricSagTable(setting.ttimen, CardNumber, BordNumber);
                    //Create_ElectricSwellTable(setting.ttimen, CardNumber, BordNumber);
                    //Create_ElectricAlarmTable(setting.ttimen, CardNumber, BordNumber);
                    using (IDbConnection connection = new SqlConnection(SqlDB))
                    {
                        string sql = $"USE [EwatchIntegrate_{setting.ttimen:yyyyMM}] IF NOT EXISTS (SELECT * FROM  ElectricSag_{CardNumber}{BordNumber} WHERE ttime = @ttime AND DeviceNumber = @DeviceNumber)" +
                            $"INSERT INTO ElectricSag_{CardNumber}{BordNumber} (" +
                            $"ttime" +
                            $",ttimen" +
                            $",DeviceNumber" +
                            $",sagCycle" +
                            $",sagData" +
                            $",sagPhase" +
                            $",sagBegin" +
                            $",sagEnd" +
                            $") VALUES (" +
                            $"@ttime" +
                            $",@ttimen" +
                            $",@DeviceNumber" +
                            $",@sagCycle" +
                            $",@sagData" +
                            $",@sagPhase" +
                            $",@sagBegin" +
                            $",@sagEnd)";
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
        /// 新增電表驟升歷史資訊
        /// </summary>
        /// <param name="setting"></param>
        /// <param name="CardNumber"></param>
        /// <param name="BordNumber"></param>
        /// <returns></returns>
        public ResponseTypeEnum ElectricSwell_Add(ElectricSwell_Log setting, string CardNumber, string BordNumber)
        {
            try
            {
                if (CardBord_Search(CardNumber, BordNumber))
                {
                    Create_DataBase(setting.ttimen);
                    //Create_ElectricTable(setting.ttimen, CardNumber, BordNumber);
                    //Create_ElectricSagTable(setting.ttimen, CardNumber, BordNumber);
                    Create_ElectricSwellTable(setting.ttimen, CardNumber, BordNumber);
                    //Create_ElectricAlarmTable(setting.ttimen, CardNumber, BordNumber);
                    using (IDbConnection connection = new SqlConnection(SqlDB))
                    {
                        string sql = $"USE [EwatchIntegrate_{setting.ttimen:yyyyMM}] IF NOT EXISTS (SELECT * FROM  ElectricSwell_{CardNumber}{BordNumber} WHERE ttime = @ttime AND DeviceNumber = @DeviceNumber)" +
                            $"INSERT INTO ElectricSwell_{CardNumber}{BordNumber} (" +
                            $"ttime" +
                            $",ttimen" +
                            $",DeviceNumber" +
                            $",swellCycle" +
                            $",swellData" +
                            $",swellPhase" +
                            $",swellBegin" +
                            $",swellEnd" +
                            $") VALUES (" +
                            $"@ttime" +
                            $",@ttimen" +
                            $",@DeviceNumber" +
                            $",@swellCycle" +
                            $",@swellData" +
                            $",@swellPhase" +
                            $",@swellBegin" +
                            $",@swellEnd)";
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
        /// 新增電表驟升歷史資訊
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
                        string sql = $"USE [EwatchIntegrate_{setting.ttimen:yyyyMM}] IF NOT EXISTS (SELECT * FROM  ElectricAlarm_{CardNumber}{BordNumber} WHERE ttime = @ttime AND DeviceNumber = @DeviceNumber)" +
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
    }
}
