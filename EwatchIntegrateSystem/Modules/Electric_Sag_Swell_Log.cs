namespace EwatchIntegrateSystem.Modules
{
    /// <summary>
    /// 驟降驟升資訊
    /// </summary>
    public class Electric_Sag_Swell_Log
    {
        /// <summary>
        /// 時間字串(yyyyMMddHHmmss)
        /// </summary>
        public string ttime { get; set; }
        /// <summary>
        /// 時間
        /// </summary>
        public DateTime ttimen { get; set; }
        /// <summary>
        /// 設備編號
        /// </summary>
        public int DeviceNumber { get; set; }
        /// <summary>
        /// 驟升驟降字串
        /// </summary>
        public string Sag_Swell_Type { get; set; }
        /// <summary>
        /// 電壓驟升之持續cycle數
        /// </summary>
        public int Cycle { get; set; }
        /// <summary>
        /// 電壓驟升發生時之百分比(通訊數值除100，資料庫讀取不用)
        /// </summary>
        public decimal Data { get; set; }
        /// <summary>
        /// 相電壓
        /// </summary>
        public string Phase { get; set; }
        /// <summary>
        /// 開始發生日期與時間(yy/MM/dd HH:mm:ss)
        /// </summary>
        public string Sag_Swell_Begin { get; set; }
        /// <summary>
        /// 結束發生日期與時間(yy/MM/dd HH:mm:ss)
        /// </summary>
        public string Sag_Swell_End { get; set; }
    }
}
