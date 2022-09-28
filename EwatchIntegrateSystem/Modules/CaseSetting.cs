using System.ComponentModel.DataAnnotations;

namespace EwatchIntegrateSystem.Modules
{
    /// <summary>
    /// 案場資料表
    /// </summary>
    public class CaseSetting
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
        /// 案場名稱
        /// </summary>
        [StringLength(50, ErrorMessage = "字串最多50個字")]
        public string? CaseName { get; set; }
    }
}
