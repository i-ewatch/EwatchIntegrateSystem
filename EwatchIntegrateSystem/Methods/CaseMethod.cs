using Dapper;
using EwatchIntegrateSystem.Enums;
using EwatchIntegrateSystem.Modules;
using System.Data;
using System.Data.SqlClient;

namespace EwatchIntegrateSystem.Methods
{
    /// <summary>
    /// 案件資料庫方法
    /// </summary>
    public class CaseMethod
    {
        private string SqlDB { get; set; }
        /// <summary>
        /// 錯誤資訊
        /// </summary>
        public string ErrorResponse { get; set; } = "";
        /// <summary>
        /// 案件資料庫方法初始化
        /// </summary>
        /// <param name="sqlDB"></param>
        public CaseMethod(string sqlDB)
        {
            SqlDB = sqlDB;
        }
        /// <summary>
        /// 查詢全部案場卡版號
        /// </summary>
        /// <returns></returns>
        public List<CaseSetting> Case_Search()
        {
            List<CaseSetting> setting = new List<CaseSetting>();
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = "SELECT * FROM CaseSetting";
                    setting = connection.Query<CaseSetting>(sql).ToList();
                }
                return setting;
            }
            catch (Exception)
            {
                return setting;
            }
        }
        /// <summary>
        /// 新增案場卡版號
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public ResponseTypeEnum Case_Add(CaseSetting setting)
        {
            try
            {
                List<CaseSetting> caseSettings = Case_Search();
                CaseSetting? caseSetting = caseSettings.SingleOrDefault(g => g.CardNumber == setting.CardNumber && g.BordNumber == setting.BordNumber);
                if (caseSetting == null)
                {
                    using (IDbConnection connection = new SqlConnection(SqlDB))
                    {
                        string sql = "INSERT INTO CaseSetting (CardNumber,BordNumber,CaseName) VALUES (@CardNumber,@BordNumber,@CaseName)";
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
        /// 修改案場名稱
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public ResponseTypeEnum Case_Update(CaseSetting setting)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = "UPDATE CaseSetting SET CaseName = @CaseName WHERE CardNumber = @CardNumber AND BordNumber = @BordNumber";
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
        /// 刪除案場
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        public ResponseTypeEnum Case_Delete(CaseSetting setting)
        {
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = "DELETE FROM CaseSetting WHERE CardNumber = @CardNumber AND BordNumber = @BordNumber";
                    string devicesql = "DELETE FROM DeviceSetting WHERE CardNumber = @CardNumber AND BordNumber = @BordNumber";
                    var Index = connection.Execute(sql, setting);
                    Index += connection.Execute(sql, devicesql);
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
