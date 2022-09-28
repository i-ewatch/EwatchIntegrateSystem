using System.ComponentModel.DataAnnotations;

namespace EwatchIntegrateSystem.Modules
{
    /// <summary>
    /// 電表告警資訊
    /// </summary>
    public class ElectricAlarm_Log
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
        /// 警報項目
        /// </summary>
        public string? AlarmItem { get; set; }
        /// <summary>
        /// 警報發生或解除時之百分比(通訊數值除100，資料庫讀取不用)
        /// </summary>
        public decimal AlarmData { get; set; }
        /// <summary>
        /// 發生日期與時間(yy/MM/dd HH:mm:ss)
        /// </summary>
        public string? AlarmTime { get; set; }
    }
}
