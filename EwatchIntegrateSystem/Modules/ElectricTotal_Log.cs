using System.ComponentModel.DataAnnotations;

namespace EwatchIntegrateSystem.Modules
{
    /// <summary>
    /// 電表累積量資訊
    /// </summary>
    public class ElectricTotal_Log
    {
        /// <summary>
        /// 時間字串(yyyyMMddHHmm)
        /// </summary>
        [StringLength(14, ErrorMessage = "字串最多14個字")]
        public string? ttime { get; set; }
        /// <summary>
        /// 時間
        /// </summary>
        public DateTime ttimen { get; set; }
        /// <summary>
        /// 卡號
        /// </summary>
        [StringLength(6, ErrorMessage = "字串最多6個字")]
        public string CardNumber { get; set; } = "";
        /// <summary>
        /// 版號
        /// </summary>
        [StringLength(2, ErrorMessage = "字串最多2個字")]
        public string BordNumber { get; set; } = "";
        /// <summary>
        /// 設備編號
        /// </summary>
        public int DeviceNumber { get; set; }
        /// <summary>
        /// 起始累積量1
        /// </summary>
        public decimal Start1 { get; set; }
        /// <summary>
        /// 結束累積量1
        /// </summary>
        public decimal End1 { get; set; }
        /// <summary>
        /// 起始累積量2
        /// </summary>
        public decimal Start2 { get; set; }
        /// <summary>
        /// 結束累積量2
        /// </summary>
        public decimal End2 { get; set; }
        /// <summary>
        /// 當日總累積量
        /// </summary>
        public decimal Total { get; set; }
    }
}
