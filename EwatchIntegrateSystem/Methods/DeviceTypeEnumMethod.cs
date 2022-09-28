using Dapper;
using EwatchIntegrateSystem.Enums;
using EwatchIntegrateSystem.Modules;
using System.Data;
using System.Data.SqlClient;

namespace EwatchIntegrateSystem.Methods
{
    /// <summary>
    /// 設備類型資料庫方法
    /// </summary>
    public class DeviceTypeEnumMethod
    {
        private string SqlDB { get; set; }
        /// <summary>
        /// 錯誤資訊
        /// </summary>
        public string ErrorResponse { get; set; } = "";
        /// <summary>
        /// 設備類型資料庫方法初始化
        /// </summary>
        /// <param name="sqlDB"></param>
        public DeviceTypeEnumMethod(string sqlDB)
        {
            SqlDB = sqlDB;
        }
        /// <summary>
        /// 查詢全部設備類型
        /// </summary>
        /// <returns></returns>
        public List<DeviceTypeEnumSetting> DeviceType_Search()
        {
            List<DeviceTypeEnumSetting> deviceTypeEnumSettings = new List<DeviceTypeEnumSetting>();
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = "SELECT * FROM DeviceTypeEnumSetting";
                    deviceTypeEnumSettings = connection.Query<DeviceTypeEnumSetting>(sql).ToList();
                }
                return deviceTypeEnumSettings;
            }
            catch (Exception ex)
            {
                ErrorResponse = ex.Message.ToString();
                return deviceTypeEnumSettings;
            }
        }
        /// <summary>
        /// 新增設備類型
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public ResponseTypeEnum DeviceType_Add(DeviceTypeEnumSetting setting)
        {
            try
            {
                List<DeviceTypeEnumSetting> deviceTypeEnumSettings = DeviceType_Search();
                DeviceTypeEnumSetting? deviceTypeEnumSetting = deviceTypeEnumSettings.SingleOrDefault(g => g.DeviceTypeEnum == setting.DeviceTypeEnum);
                if (deviceTypeEnumSetting == null)
                {

                    using (IDbConnection connection = new SqlConnection(SqlDB))
                    {
                        string sql = "INSERT INTO DeviceTypeEnumSetting (DeviceTypeEnum,DeviceTypeName) VALUES (@DeviceTypeEnum,@DeviceTypeName)";
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
        /// 修改設備類型
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public ResponseTypeEnum DeviceType_Update(DeviceTypeEnumSetting setting)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = "UPDATE DeviceTypeEnumSetting SET DeviceTypeName = @DeviceTypeName WHERE DeviceTypeEnum = @DeviceTypeEnum";
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
        /// 刪除設備類型
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public ResponseTypeEnum DeviceType_Delete(DeviceTypeEnumSetting setting)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = "DELETE FROM DeviceTypeEnumSetting WHERE DeviceTypeEnum = @DeviceTypeEnum";
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
