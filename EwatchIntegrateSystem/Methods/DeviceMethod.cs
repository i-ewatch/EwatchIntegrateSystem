using Dapper;
using EwatchIntegrateSystem.Enums;
using EwatchIntegrateSystem.Modules;
using System.Data;
using System.Data.SqlClient;

namespace EwatchIntegrateSystem.Methods
{
    /// <summary>
    /// 設備資料庫方法
    /// </summary>
    public class DeviceMethod
    {
        private string SqlDB { get; set; }
        /// <summary>
        /// 錯誤資訊
        /// </summary>
        public string ErrorResponse { get; set; } = "";
        /// <summary>
        /// 設備資料庫方法初始化
        /// </summary>
        /// <param name="sqlDB"></param>
        public DeviceMethod(string sqlDB)
        {
            SqlDB = sqlDB;
        }
        /// <summary>
        /// 查詢全部設備
        /// </summary>
        /// <returns></returns>
        public List<DeviceSetting> Device_Search()
        {
            List<DeviceSetting> setting = new List<DeviceSetting>();
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = "SELECT * FROM DeviceSetting";
                    setting = connection.Query<DeviceSetting>(sql).ToList();
                }
                return setting;
            }
            catch (Exception)
            {
                return setting;
            }
        }
        /// <summary>
        /// 查詢特定卡版號設備
        /// </summary>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        public List<DeviceSetting> Device_Search(string CardNumber, string BordNumber)
        {
            List<DeviceSetting> setting = new List<DeviceSetting>();
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = "SELECT * FROM DeviceSetting WHERE CardNumber = @CardNumber AND BordNumber = @BordNumber";
                    setting = connection.Query<DeviceSetting>(sql, new { CardNumber, BordNumber }).ToList();
                }
                return setting;
            }
            catch (Exception)
            {
                return setting;
            }
        }
        /// <summary>
        /// 新增特定卡版號設備(單)
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public ResponseTypeEnum Device_Single_Add(DeviceSetting setting)
        {
            try
            {
                List<DeviceSetting> deviceSettings = Device_Search(setting.CardNumber, setting.BordNumber);
                DeviceSetting? deviceSetting = deviceSettings.SingleOrDefault(g => g.DeviceNumber == setting.DeviceNumber);
                if (deviceSetting == null)
                {
                    using (IDbConnection connection = new SqlConnection(SqlDB))
                    {
                        string sql = "INSERT INTO DeviceSetting (CardNumber,BordNumber,DeviceNumber,DeviceName,DeviceTypeEnum) VALUES (@CardNumber,@BordNumber,@DeviceNumber,@DeviceName,@DeviceTypeEnum)";
                        var Index = connection.Execute(sql, setting);
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
                    return ResponseTypeEnum.Repeat;
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return ResponseTypeEnum.Error;
            }
        }
        /// <summary>
        /// 新增特定卡版號設備(多)
        /// </summary>
        /// <param name="settings"></param>
        /// <returns></returns>
        public ResponseTypeEnum Device_Multiple_Add(List<DeviceSetting> settings)
        {
            try
            {
                List<DeviceSetting> deviceSettings = Device_Search(settings[0].CardNumber, settings[0].BordNumber);
                DeviceSetting? deviceSetting = deviceSettings.SingleOrDefault(g => g.DeviceNumber == settings[0].DeviceNumber);
                if (deviceSetting == null)
                {
                    using (IDbConnection connection = new SqlConnection(SqlDB))
                    {
                        string sql = "INSERT INTO DeviceSetting (CardNumber,BordNumber,DeviceNumber,DeviceName,DeviceTypeEnum) VALUES (@CardNumber,@BordNumber,@DeviceNumber,@DeviceName,@DeviceTypeEnum)";
                        var Index = connection.Execute(sql, settings);
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
                    return ResponseTypeEnum.Repeat;
                }
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return ResponseTypeEnum.Error;
            }
        }
        /// <summary>
        /// 修改特定卡版號設備(單)
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public ResponseTypeEnum Device_Single_Update(DeviceSetting setting)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = "UPDATE DeviceSetting SET DeviceName = @DeviceName,DeviceTypeEnum = @DeviceTypeEnum WHERE CardNumber = @CardNumber AND BordNumber = @BordNumber AND DeviceNumber = @DeviceNumber";
                    var Index = connection.Execute(sql, setting);
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
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return ResponseTypeEnum.Error;
            }
        }
        /// <summary>
        /// 修改特定卡版號設備(多)
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public ResponseTypeEnum Device_Multiple_Update(List<DeviceSetting> setting)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = "UPDATE DeviceSetting SET DeviceName = @DeviceName,DeviceTypeEnum = @DeviceTypeEnum WHERE CardNumber = @CardNumber AND BordNumber = @BordNumber AND DeviceNumber = @DeviceNumber";
                    var Index = connection.Execute(sql, setting);
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
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return ResponseTypeEnum.Error;
            }
        }
        /// <summary>
        /// 刪除特定卡版號設備(單)
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public ResponseTypeEnum Device_Single_Delete(DeviceSetting setting)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = "DELETE FROM DeviceSetting WHERE CardNumber = @CardNumber AND BordNumber = @BordNumber AND DeviceNumber = @DeviceNumber";
                    var Index = connection.Execute(sql, setting);
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
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return ResponseTypeEnum.Error;
            }
        }
        /// <summary>
        /// 刪除特定卡版號設備(多)
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public ResponseTypeEnum Device_Multiple_Delete(List<DeviceSetting> setting)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = "DELETE FROM DeviceSetting WHERE CardNumber = @CardNumber AND BordNumber = @BordNumber AND DeviceNumber = @DeviceNumber";
                    var Index = connection.Execute(sql, setting);
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
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return ResponseTypeEnum.Error;
            }
        }
    }
}
