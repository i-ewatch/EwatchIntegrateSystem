using System.ComponentModel.DataAnnotations;

namespace EwatchIntegrateSystem.Modules
{
    /// <summary>
    /// 設備類型資料表
    /// </summary>
    public class DeviceTypeEnumSetting
    {
        /// <summary>
        /// 設備類型編號
        /// </summary>
        public int DeviceTypeEnum { get; set; }
        /// <summary>
        /// 設備名稱
        /// </summary>
        [StringLength(10, ErrorMessage = "字串最多10個字")]
        public string? DeviceTypeName { get; set; }
    }
}
