using System.ComponentModel.DataAnnotations;

namespace EwatchIntegrateSystem.Modules
{
    /// <summary>
    /// 電表驟升電壓資訊
    /// </summary>
    public class ElectricSwell_Log
    {
        /// <summary>
        /// 時間字串(yyyyMMddHHmmss)
        /// </summary>
        [StringLength(16, ErrorMessage = "字串最多16個字")]
        public string? ttime { get; set; }
        /// <summary>
        /// 時間
        /// </summary>
        public DateTime ttimen { get; set; }
        /// <summary>
        /// 設備編號
        /// </summary>
        public int DeviceNumber { get; set; }
        /// <summary>
        /// 電壓驟升之持續cycle數
        /// </summary>
        public int swellCycle { get; set; }
        /// <summary>
        /// 電壓驟升發生時之百分比(通訊數值除100，資料庫讀取不用)
        /// </summary>
        public decimal swellData { get; set; }
        /// <summary>
        /// 相電壓
        /// </summary>
        [StringLength(2, ErrorMessage = "字串最多2個字")]
        public string? swellPhase { get; set; }
        /// <summary>
        /// 開始發生日期與時間(yy/MM/dd HH:mm:ss)
        /// </summary>
        public string? swellBegin { get; set; }
        /// <summary>
        /// 結束發生日期與時間(yy/MM/dd HH:mm:ss)
        /// </summary>
        public string? swellEnd { get; set; }
    }
}
