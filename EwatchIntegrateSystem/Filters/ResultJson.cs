namespace EwatchIntegrateSystem.Filters
{
    /// <summary>
    /// 統一回應格式
    /// </summary>
    public class ResultJson
    {
        /// <summary>
        /// 數值資訊
        /// </summary>
        public object? Data { get; set; } = "Error";
        /// <summary>
        /// Code碼
        /// </summary>
        public int HttpCode { get; set; } = 401;
        /// <summary>
        /// 錯誤訊息
        /// </summary>
        public string? Error { get; set; }
    }
}
