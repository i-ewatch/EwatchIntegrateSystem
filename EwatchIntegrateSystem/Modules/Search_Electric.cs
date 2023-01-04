namespace EwatchIntegrateSystem.Modules
{
    /// <summary>
    /// 電表查詢物件
    /// </summary>
    public class Search_Electric
    {
        /// <summary>
        /// 開始時間
        /// </summary>
        public DateTime StartTime { get; set; }
        /// <summary>
        /// 結束時間
        /// </summary>
        public DateTime EndTime { get; set; }
        /// <summary>
        /// 卡號
        /// </summary>
        public string CardNumber { get; set; } = string.Empty;
        /// <summary>
        /// 版號
        /// </summary>
        public string BordNumber { get; set; } = string.Empty;
        /// <summary>
        /// 設備名稱
        /// </summary>
        public string DeviceName { get; set; } = string.Empty;
        /// <summary>
        /// 設備編號
        /// </summary>
        public int DeviceNumber { get; set; }
        /// <summary>
        /// 報表查詢類型
        /// </summary>
        public int Search_Type { get; set; }
        /// <summary>
        /// 查詢項目
        /// </summary>
        public int Item_Type { get; set; }
    }
}
