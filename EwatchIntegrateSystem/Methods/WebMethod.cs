using Dapper;
using EwatchIntegrateSystem.Modules;
using System.Data;
using System.Data.SqlClient;

namespace EwatchIntegrateSystem.Methods
{
    /// <summary>
    /// 即時資訊資料庫方法
    /// </summary>
    public class WebMethod
    {
        private string SqlDB { get; set; }
        /// <summary>
        /// 錯誤資訊
        /// </summary>
        public string ErrorResponse { get; set; } = "";
        /// <summary>
        /// 即時資訊資料庫方法初始化
        /// </summary>
        /// <param name="sqlDB"></param>
        public WebMethod(string sqlDB)
        {
            SqlDB = sqlDB;
        }
        /// <summary>
        /// 查詢全部電表即時資訊
        /// </summary>
        /// <returns></returns>
        public List<ElectricWeb> ElectricWeb_Search()
        {
            List<ElectricWeb> setting = new List<ElectricWeb>();
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = "SELECT * FROM ElectricWeb";
                    setting = connection.Query<ElectricWeb>(sql).ToList();
                }
                return setting;
            }
            catch (Exception)
            {
                return setting;
            }
        }
        /// <summary>
        /// 查詢特定卡版號電表即時資訊
        /// </summary>
        /// <param name="CardNumber"></param>
        /// <param name="BordNumber"></param>
        /// <returns></returns>
        public List<ElectricWeb> ElectricWeb_Search(string CardNumber, string BordNumber)
        {
            List<ElectricWeb> setting = new List<ElectricWeb>();
            try
            {
                using (IDbConnection connection = new SqlConnection(SqlDB))
                {
                    string sql = "SELECT * FROM ElectricWeb WHERE CardNumber = @CardNumber AND BordNumber = @BordNumber";
                    setting = connection.Query<ElectricWeb>(sql, new { CardNumber, BordNumber }).ToList();
                }
                return setting;
            }
            catch (Exception)
            {
                return setting;
            }
        }
    }
}
