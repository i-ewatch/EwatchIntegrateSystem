using System.ComponentModel.DataAnnotations;

namespace EwatchIntegrateSystem.Modules
{
    /// <summary>
    /// 設備資料表
    /// </summary>
    public class DeviceSetting
    {
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
        /// 設備名稱
        /// </summary>
        [StringLength(20, ErrorMessage = "字串最多20個字")]
        public string? DeviceName { get; set; }
        /// <summary>
        /// 設備類型
        /// </summary>
        public int DeviceTypeEnum { get; set; }
    }
}
