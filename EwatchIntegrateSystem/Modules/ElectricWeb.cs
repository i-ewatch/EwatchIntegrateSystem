using System.ComponentModel.DataAnnotations;

namespace EwatchIntegrateSystem.Modules
{
    /// <summary>
    /// 網頁電表資訊
    /// </summary>
    public class ElectricWeb
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
        /// a相電壓
        /// </summary>
        public decimal Va { get; set; }
        /// <summary>
        /// b相電壓
        /// </summary>
        public decimal Vb { get; set; }
        /// <summary>
        /// c相電壓
        /// </summary>
        public decimal Vc { get; set; }
        /// <summary>
        /// 平均相電壓
        /// </summary>
        public decimal Vnavg { get; set; }
        /// <summary>
        /// a線電壓
        /// </summary>
        public decimal Vab { get; set; }
        /// <summary>
        /// b線電壓
        /// </summary>
        public decimal Vbc { get; set; }
        /// <summary>
        /// c線電壓
        /// </summary>
        public decimal Vca { get; set; }
        /// <summary>
        /// 平均線電壓
        /// </summary>
        public decimal Vlavg { get; set; }
        /// <summary>
        /// a相電流
        /// </summary>
        public decimal Ia { get; set; }
        /// <summary>
        /// b相電流
        /// </summary>
        public decimal Ib { get; set; }
        /// <summary>
        /// c相電流
        /// </summary>
        public decimal Ic { get; set; }
        /// <summary>
        /// 平均相電流
        /// </summary>
        public decimal Iavg { get; set; }
        /// <summary>
        /// 頻率
        /// </summary>
        public decimal Freq { get; set; }
        /// <summary>
        /// a有效功率
        /// </summary>
        public decimal Kwa { get; set; }
        /// <summary>
        /// b有效功率
        /// </summary>
        public decimal Kwb { get; set; }
        /// <summary>
        /// c有效功率
        /// </summary>
        public decimal Kwc { get; set; }
        /// <summary>
        /// 總有效功率
        /// </summary>
        public decimal Kwtot { get; set; }
        /// <summary>
        /// a無效功率
        /// </summary>
        public decimal KVARa { get; set; }
        /// <summary>
        /// b無效功率
        /// </summary>
        public decimal KVARb { get; set; }
        /// <summary>
        /// c無效功率
        /// </summary>
        public decimal KVARc { get; set; }
        /// <summary>
        /// 總無效功率
        /// </summary>
        public decimal KVARtot { get; set; }
        /// <summary>
        /// a視在功率
        /// </summary>
        public decimal KVAa { get; set; }
        /// <summary>
        /// b視在功率
        /// </summary>
        public decimal KVAb { get; set; }
        /// <summary>
        /// c視在功率
        /// </summary>
        public decimal KVAc { get; set; }
        /// <summary>
        /// 總視在功率
        /// </summary>
        public decimal KVAtot { get; set; }
        /// <summary>
        /// a功率因數
        /// </summary>
        public decimal PFa { get; set; }
        /// <summary>
        /// b功率因數
        /// </summary>
        public decimal PFb { get; set; }
        /// <summary>
        /// c功率因數
        /// </summary>
        public decimal PFc { get; set; }
        /// <summary>
        /// 平均功率因數
        /// </summary>
        public decimal PFavg { get; set; }
        /// <summary>
        /// a相電壓相位角
        /// </summary>
        public decimal PhaseAngleVa { get; set; }
        /// <summary>
        /// b相電壓相位角
        /// </summary>
        public decimal PhaseAngleVb { get; set; }
        /// <summary>
        /// c相電壓相位角
        /// </summary>
        public decimal PhaseAngleVc { get; set; }
        /// <summary>
        /// a相電流相位角
        /// </summary>
        public decimal PhaseAngleIa { get; set; }
        /// <summary>
        /// b相電流相位角
        /// </summary>
        public decimal PhaseAngleIb { get; set; }
        /// <summary>
        /// c相電流相位角
        /// </summary>
        public decimal PhaseAngleIc { get; set; }
        /// <summary>
        /// 累積功率
        /// </summary>
        public decimal KWH { get; set; }
    }
}
