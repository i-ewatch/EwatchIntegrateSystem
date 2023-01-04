using System.ComponentModel.DataAnnotations;

namespace EwatchIntegrateSystem.Modules
{
    /// <summary>
    /// 電表歷史資訊
    /// </summary>
    public class ElectricInformation_Log
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
        /// 設備編號
        /// </summary>
        public int DeviceNumber { get; set; }
        /// <summary>
        /// 設備迴路
        /// </summary>
        public int Loop { get; set; }
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
        /// 用電百分比
        /// </summary>
        public decimal KwPercent { get; set; }
        #region RV諧波
        /// <summary>
        /// RV諧波
        /// </summary>
        public decimal RV_HD { get; set; }
        /// <summary>
        /// RV諧波_1
        /// </summary>
        public decimal RV_HD1 { get; set; }
        /// <summary>
        /// RV諧波_2
        /// </summary>
        public decimal RV_HD2 { get; set; }
        /// <summary>
        /// RV諧波_3
        /// </summary>
        public decimal RV_HD3 { get; set; }
        /// <summary>
        /// RV諧波_4
        /// </summary>
        public decimal RV_HD4 { get; set; }
        /// <summary>
        /// RV諧波_5
        /// </summary>
        public decimal RV_HD5 { get; set; }
        /// <summary>
        /// RV諧波_6
        /// </summary>
        public decimal RV_HD6 { get; set; }
        /// <summary>
        /// RV諧波_7
        /// </summary>
        public decimal RV_HD7 { get; set; }
        /// <summary>
        /// RV諧波_8
        /// </summary>
        public decimal RV_HD8 { get; set; }
        /// <summary>
        /// RV諧波_9
        /// </summary>
        public decimal RV_HD9 { get; set; }
        /// <summary>
        /// RV諧波_10
        /// </summary>
        public decimal RV_HD10 { get; set; }
        /// <summary>
        /// RV諧波_11
        /// </summary>
        public decimal RV_HD11 { get; set; }
        /// <summary>
        /// RV諧波_12
        /// </summary>
        public decimal RV_HD12 { get; set; }
        /// <summary>
        /// RV諧波_13
        /// </summary>
        public decimal RV_HD13 { get; set; }
        /// <summary>
        /// RV諧波_14
        /// </summary>
        public decimal RV_HD14 { get; set; }
        /// <summary>
        /// RV諧波_15
        /// </summary>
        public decimal RV_HD15 { get; set; }
        /// <summary>
        /// RV諧波_16
        /// </summary>
        public decimal RV_HD16 { get; set; }
        /// <summary>
        /// RV諧波_17
        /// </summary>
        public decimal RV_HD17 { get; set; }
        /// <summary>
        /// RV諧波_18
        /// </summary>
        public decimal RV_HD18 { get; set; }
        /// <summary>
        /// RV諧波_19
        /// </summary>
        public decimal RV_HD19 { get; set; }
        /// <summary>
        /// RV諧波_20
        /// </summary>
        public decimal RV_HD20 { get; set; }
        /// <summary>
        /// RV諧波_21
        /// </summary>
        public decimal RV_HD21 { get; set; }
        /// <summary>
        /// RV諧波_22
        /// </summary>
        public decimal RV_HD22 { get; set; }
        /// <summary>
        /// RV諧波_23
        /// </summary>
        public decimal RV_HD23 { get; set; }
        /// <summary>
        /// RV諧波_24
        /// </summary>
        public decimal RV_HD24 { get; set; }
        /// <summary>
        /// RV諧波_25
        /// </summary>
        public decimal RV_HD25 { get; set; }
        /// <summary>
        /// RV諧波_26
        /// </summary>
        public decimal RV_HD26 { get; set; }
        /// <summary>
        /// RV諧波_27
        /// </summary>
        public decimal RV_HD27 { get; set; }
        /// <summary>
        /// RV諧波_28
        /// </summary>
        public decimal RV_HD28 { get; set; }
        /// <summary>
        /// RV諧波_29
        /// </summary>
        public decimal RV_HD29 { get; set; }
        /// <summary>
        /// RV諧波_30
        /// </summary>
        public decimal RV_HD30 { get; set; }
        /// <summary>
        /// RV諧波_31
        /// </summary>
        public decimal RV_HD31 { get; set; }
        #endregion
        #region SV諧波
        /// <summary>
        /// SV諧波
        /// </summary>
        public decimal SV_HD { get; set; }
        /// <summary>
        /// SV諧波_1
        /// </summary>
        public decimal SV_HD1 { get; set; }
        /// <summary>
        /// SV諧波_2
        /// </summary>
        public decimal SV_HD2 { get; set; }
        /// <summary>
        /// SV諧波_3
        /// </summary>
        public decimal SV_HD3 { get; set; }
        /// <summary>
        /// SV諧波_4
        /// </summary>
        public decimal SV_HD4 { get; set; }
        /// <summary>
        /// SV諧波_5
        /// </summary>
        public decimal SV_HD5 { get; set; }
        /// <summary>
        /// SV諧波_6
        /// </summary>
        public decimal SV_HD6 { get; set; }
        /// <summary>
        /// SV諧波_7
        /// </summary>
        public decimal SV_HD7 { get; set; }
        /// <summary>
        /// SV諧波_8
        /// </summary>
        public decimal SV_HD8 { get; set; }
        /// <summary>
        /// SV諧波_9
        /// </summary>
        public decimal SV_HD9 { get; set; }
        /// <summary>
        /// SV諧波_10
        /// </summary>
        public decimal SV_HD10 { get; set; }
        /// <summary>
        /// SV諧波_11
        /// </summary>
        public decimal SV_HD11 { get; set; }
        /// <summary>
        /// SV諧波_12
        /// </summary>
        public decimal SV_HD12 { get; set; }
        /// <summary>
        /// SV諧波_13
        /// </summary>
        public decimal SV_HD13 { get; set; }
        /// <summary>
        /// SV諧波_14
        /// </summary>
        public decimal SV_HD14 { get; set; }
        /// <summary>
        /// SV諧波_15
        /// </summary>
        public decimal SV_HD15 { get; set; }
        /// <summary>
        /// SV諧波_16
        /// </summary>
        public decimal SV_HD16 { get; set; }
        /// <summary>
        /// SV諧波_17
        /// </summary>
        public decimal SV_HD17 { get; set; }
        /// <summary>
        /// SV諧波_18
        /// </summary>
        public decimal SV_HD18 { get; set; }
        /// <summary>
        /// SV諧波_19
        /// </summary>
        public decimal SV_HD19 { get; set; }
        /// <summary>
        /// SV諧波_20
        /// </summary>
        public decimal SV_HD20 { get; set; }
        /// <summary>
        /// SV諧波_21
        /// </summary>
        public decimal SV_HD21 { get; set; }
        /// <summary>
        /// SV諧波_22
        /// </summary>
        public decimal SV_HD22 { get; set; }
        /// <summary>
        /// SV諧波_23
        /// </summary>
        public decimal SV_HD23 { get; set; }
        /// <summary>
        /// SV諧波_24
        /// </summary>
        public decimal SV_HD24 { get; set; }
        /// <summary>
        /// SV諧波_25
        /// </summary>
        public decimal SV_HD25 { get; set; }
        /// <summary>
        /// SV諧波_26
        /// </summary>
        public decimal SV_HD26 { get; set; }
        /// <summary>
        /// SV諧波_27
        /// </summary>
        public decimal SV_HD27 { get; set; }
        /// <summary>
        /// SV諧波_28
        /// </summary>
        public decimal SV_HD28 { get; set; }
        /// <summary>
        /// SV諧波_29
        /// </summary>
        public decimal SV_HD29 { get; set; }
        /// <summary>
        /// SV諧波_30
        /// </summary>
        public decimal SV_HD30 { get; set; }
        /// <summary>
        /// SV諧波_31
        /// </summary>
        public decimal SV_HD31 { get; set; }
        #endregion
        #region TV諧波
        /// <summary>
        /// TV諧波
        /// </summary>
        public decimal TV_HD { get; set; }
        /// <summary>
        /// TV諧波_1
        /// </summary>
        public decimal TV_HD1 { get; set; }
        /// <summary>
        /// TV諧波_2
        /// </summary>
        public decimal TV_HD2 { get; set; }
        /// <summary>
        /// TV諧波_3
        /// </summary>
        public decimal TV_HD3 { get; set; }
        /// <summary>
        /// TV諧波_4
        /// </summary>
        public decimal TV_HD4 { get; set; }
        /// <summary>
        /// TV諧波_5
        /// </summary>
        public decimal TV_HD5 { get; set; }
        /// <summary>
        /// TV諧波_6
        /// </summary>
        public decimal TV_HD6 { get; set; }
        /// <summary>
        /// TV諧波_7
        /// </summary>
        public decimal TV_HD7 { get; set; }
        /// <summary>
        /// TV諧波_8
        /// </summary>
        public decimal TV_HD8 { get; set; }
        /// <summary>
        /// TV諧波_9
        /// </summary>
        public decimal TV_HD9 { get; set; }
        /// <summary>
        /// TV諧波_10
        /// </summary>
        public decimal TV_HD10 { get; set; }
        /// <summary>
        /// TV諧波_11
        /// </summary>
        public decimal TV_HD11 { get; set; }
        /// <summary>
        /// TV諧波_12
        /// </summary>
        public decimal TV_HD12 { get; set; }
        /// <summary>
        /// TV諧波_13
        /// </summary>
        public decimal TV_HD13 { get; set; }
        /// <summary>
        /// TV諧波_14
        /// </summary>
        public decimal TV_HD14 { get; set; }
        /// <summary>
        /// TV諧波_15
        /// </summary>
        public decimal TV_HD15 { get; set; }
        /// <summary>
        /// TV諧波_16
        /// </summary>
        public decimal TV_HD16 { get; set; }
        /// <summary>
        /// TV諧波_17
        /// </summary>
        public decimal TV_HD17 { get; set; }
        /// <summary>
        /// TV諧波_18
        /// </summary>
        public decimal TV_HD18 { get; set; }
        /// <summary>
        /// TV諧波_19
        /// </summary>
        public decimal TV_HD19 { get; set; }
        /// <summary>
        /// TV諧波_20
        /// </summary>
        public decimal TV_HD20 { get; set; }
        /// <summary>
        /// TV諧波_21
        /// </summary>
        public decimal TV_HD21 { get; set; }
        /// <summary>
        /// TV諧波_22
        /// </summary>
        public decimal TV_HD22 { get; set; }
        /// <summary>
        /// TV諧波_23
        /// </summary>
        public decimal TV_HD23 { get; set; }
        /// <summary>
        /// TV諧波_24
        /// </summary>
        public decimal TV_HD24 { get; set; }
        /// <summary>
        /// TV諧波_25
        /// </summary>
        public decimal TV_HD25 { get; set; }
        /// <summary>
        /// TV諧波_26
        /// </summary>
        public decimal TV_HD26 { get; set; }
        /// <summary>
        /// TV諧波_27
        /// </summary>
        public decimal TV_HD27 { get; set; }
        /// <summary>
        /// TV諧波_28
        /// </summary>
        public decimal TV_HD28 { get; set; }
        /// <summary>
        /// TV諧波_29
        /// </summary>
        public decimal TV_HD29 { get; set; }
        /// <summary>
        /// TV諧波_30
        /// </summary>
        public decimal TV_HD30 { get; set; }
        /// <summary>
        /// TV諧波_31
        /// </summary>
        public decimal TV_HD31 { get; set; }
        #endregion
        #region RI諧波
        /// <summary>
        /// RI諧波
        /// </summary>
        public decimal RI_HD { get; set; }
        /// <summary>
        /// RI諧波_1
        /// </summary>
        public decimal RI_HD1 { get; set; }
        /// <summary>
        /// RI諧波_2
        /// </summary>
        public decimal RI_HD2 { get; set; }
        /// <summary>
        /// RI諧波_3
        /// </summary>
        public decimal RI_HD3 { get; set; }
        /// <summary>
        /// RI諧波_4
        /// </summary>
        public decimal RI_HD4 { get; set; }
        /// <summary>
        /// RI諧波_5
        /// </summary>
        public decimal RI_HD5 { get; set; }
        /// <summary>
        /// RI諧波_6
        /// </summary>
        public decimal RI_HD6 { get; set; }
        /// <summary>
        /// RI諧波_7
        /// </summary>
        public decimal RI_HD7 { get; set; }
        /// <summary>
        /// RI諧波_8
        /// </summary>
        public decimal RI_HD8 { get; set; }
        /// <summary>
        /// RI諧波_9
        /// </summary>
        public decimal RI_HD9 { get; set; }
        /// <summary>
        /// RI諧波_10
        /// </summary>
        public decimal RI_HD10 { get; set; }
        /// <summary>
        /// RI諧波_11
        /// </summary>
        public decimal RI_HD11 { get; set; }
        /// <summary>
        /// RI諧波_12
        /// </summary>
        public decimal RI_HD12 { get; set; }
        /// <summary>
        /// RI諧波_13
        /// </summary>
        public decimal RI_HD13 { get; set; }
        /// <summary>
        /// RI諧波_14
        /// </summary>
        public decimal RI_HD14 { get; set; }
        /// <summary>
        /// RI諧波_15
        /// </summary>
        public decimal RI_HD15 { get; set; }
        /// <summary>
        /// RI諧波_16
        /// </summary>
        public decimal RI_HD16 { get; set; }
        /// <summary>
        /// RI諧波_17
        /// </summary>
        public decimal RI_HD17 { get; set; }
        /// <summary>
        /// RI諧波_18
        /// </summary>
        public decimal RI_HD18 { get; set; }
        /// <summary>
        /// RI諧波_19
        /// </summary>
        public decimal RI_HD19 { get; set; }
        /// <summary>
        /// RI諧波_20
        /// </summary>
        public decimal RI_HD20 { get; set; }
        /// <summary>
        /// RI諧波_21
        /// </summary>
        public decimal RI_HD21 { get; set; }
        /// <summary>
        /// RI諧波_22
        /// </summary>
        public decimal RI_HD22 { get; set; }
        /// <summary>
        /// RI諧波_23
        /// </summary>
        public decimal RI_HD23 { get; set; }
        /// <summary>
        /// RI諧波_24
        /// </summary>
        public decimal RI_HD24 { get; set; }
        /// <summary>
        /// RI諧波_25
        /// </summary>
        public decimal RI_HD25 { get; set; }
        /// <summary>
        /// RI諧波_26
        /// </summary>
        public decimal RI_HD26 { get; set; }
        /// <summary>
        /// RI諧波_27
        /// </summary>
        public decimal RI_HD27 { get; set; }
        /// <summary>
        /// RI諧波_28
        /// </summary>
        public decimal RI_HD28 { get; set; }
        /// <summary>
        /// RI諧波_29
        /// </summary>
        public decimal RI_HD29 { get; set; }
        /// <summary>
        /// RI諧波_30
        /// </summary>
        public decimal RI_HD30 { get; set; }
        /// <summary>
        /// RI諧波_31
        /// </summary>
        public decimal RI_HD31 { get; set; }
        #endregion
        #region SI諧波
        /// <summary>
        /// SI諧波
        /// </summary>
        public decimal SI_HD { get; set; }
        /// <summary>
        /// SI諧波_1
        /// </summary>
        public decimal SI_HD1 { get; set; }
        /// <summary>
        /// SI諧波_2
        /// </summary>
        public decimal SI_HD2 { get; set; }
        /// <summary>
        /// SI諧波_3
        /// </summary>
        public decimal SI_HD3 { get; set; }
        /// <summary>
        /// SI諧波_4
        /// </summary>
        public decimal SI_HD4 { get; set; }
        /// <summary>
        /// SI諧波_5
        /// </summary>
        public decimal SI_HD5 { get; set; }
        /// <summary>
        /// SI諧波_6
        /// </summary>
        public decimal SI_HD6 { get; set; }
        /// <summary>
        /// SI諧波_7
        /// </summary>
        public decimal SI_HD7 { get; set; }
        /// <summary>
        /// SI諧波_8
        /// </summary>
        public decimal SI_HD8 { get; set; }
        /// <summary>
        /// SI諧波_9
        /// </summary>
        public decimal SI_HD9 { get; set; }
        /// <summary>
        /// SI諧波_10
        /// </summary>
        public decimal SI_HD10 { get; set; }
        /// <summary>
        /// SI諧波_11
        /// </summary>
        public decimal SI_HD11 { get; set; }
        /// <summary>
        /// SI諧波_12
        /// </summary>
        public decimal SI_HD12 { get; set; }
        /// <summary>
        /// SI諧波_13
        /// </summary>
        public decimal SI_HD13 { get; set; }
        /// <summary>
        /// SI諧波_14
        /// </summary>
        public decimal SI_HD14 { get; set; }
        /// <summary>
        /// SI諧波_15
        /// </summary>
        public decimal SI_HD15 { get; set; }
        /// <summary>
        /// SI諧波_16
        /// </summary>
        public decimal SI_HD16 { get; set; }
        /// <summary>
        /// SI諧波_17
        /// </summary>
        public decimal SI_HD17 { get; set; }
        /// <summary>
        /// SI諧波_18
        /// </summary>
        public decimal SI_HD18 { get; set; }
        /// <summary>
        /// SI諧波_19
        /// </summary>
        public decimal SI_HD19 { get; set; }
        /// <summary>
        /// SI諧波_20
        /// </summary>
        public decimal SI_HD20 { get; set; }
        /// <summary>
        /// SI諧波_21
        /// </summary>
        public decimal SI_HD21 { get; set; }
        /// <summary>
        /// SI諧波_22
        /// </summary>
        public decimal SI_HD22 { get; set; }
        /// <summary>
        /// SI諧波_23
        /// </summary>
        public decimal SI_HD23 { get; set; }
        /// <summary>
        /// SI諧波_24
        /// </summary>
        public decimal SI_HD24 { get; set; }
        /// <summary>
        /// SI諧波_25
        /// </summary>
        public decimal SI_HD25 { get; set; }
        /// <summary>
        /// SI諧波_26
        /// </summary>
        public decimal SI_HD26 { get; set; }
        /// <summary>
        /// SI諧波_27
        /// </summary>
        public decimal SI_HD27 { get; set; }
        /// <summary>
        /// SI諧波_28
        /// </summary>
        public decimal SI_HD28 { get; set; }
        /// <summary>
        /// SI諧波_29
        /// </summary>
        public decimal SI_HD29 { get; set; }
        /// <summary>
        /// SI諧波_30
        /// </summary>
        public decimal SI_HD30 { get; set; }
        /// <summary>
        /// SI諧波_31
        /// </summary>
        public decimal SI_HD31 { get; set; }
        #endregion
        #region TI諧波
        /// <summary>
        /// TI諧波
        /// </summary>
        public decimal TI_HD { get; set; }
        /// <summary>
        /// TI諧波_1
        /// </summary>
        public decimal TI_HD1 { get; set; }
        /// <summary>
        /// TI諧波_2
        /// </summary>
        public decimal TI_HD2 { get; set; }
        /// <summary>
        /// TI諧波_3
        /// </summary>
        public decimal TI_HD3 { get; set; }
        /// <summary>
        /// TI諧波_4
        /// </summary>
        public decimal TI_HD4 { get; set; }
        /// <summary>
        /// TI諧波_5
        /// </summary>
        public decimal TI_HD5 { get; set; }
        /// <summary>
        /// TI諧波_6
        /// </summary>
        public decimal TI_HD6 { get; set; }
        /// <summary>
        /// TI諧波_7
        /// </summary>
        public decimal TI_HD7 { get; set; }
        /// <summary>
        /// TI諧波_8
        /// </summary>
        public decimal TI_HD8 { get; set; }
        /// <summary>
        /// TI諧波_9
        /// </summary>
        public decimal TI_HD9 { get; set; }
        /// <summary>
        /// TI諧波_10
        /// </summary>
        public decimal TI_HD10 { get; set; }
        /// <summary>
        /// TI諧波_11
        /// </summary>
        public decimal TI_HD11 { get; set; }
        /// <summary>
        /// TI諧波_12
        /// </summary>
        public decimal TI_HD12 { get; set; }
        /// <summary>
        /// TI諧波_13
        /// </summary>
        public decimal TI_HD13 { get; set; }
        /// <summary>
        /// TI諧波_14
        /// </summary>
        public decimal TI_HD14 { get; set; }
        /// <summary>
        /// TI諧波_15
        /// </summary>
        public decimal TI_HD15 { get; set; }
        /// <summary>
        /// TI諧波_16
        /// </summary>
        public decimal TI_HD16 { get; set; }
        /// <summary>
        /// TI諧波_17
        /// </summary>
        public decimal TI_HD17 { get; set; }
        /// <summary>
        /// TI諧波_18
        /// </summary>
        public decimal TI_HD18 { get; set; }
        /// <summary>
        /// TI諧波_19
        /// </summary>
        public decimal TI_HD19 { get; set; }
        /// <summary>
        /// TI諧波_20
        /// </summary>
        public decimal TI_HD20 { get; set; }
        /// <summary>
        /// TI諧波_21
        /// </summary>
        public decimal TI_HD21 { get; set; }
        /// <summary>
        /// TI諧波_22
        /// </summary>
        public decimal TI_HD22 { get; set; }
        /// <summary>
        /// TI諧波_23
        /// </summary>
        public decimal TI_HD23 { get; set; }
        /// <summary>
        /// TI諧波_24
        /// </summary>
        public decimal TI_HD24 { get; set; }
        /// <summary>
        /// TI諧波_25
        /// </summary>
        public decimal TI_HD25 { get; set; }
        /// <summary>
        /// TI諧波_26
        /// </summary>
        public decimal TI_HD26 { get; set; }
        /// <summary>
        /// TI諧波_27
        /// </summary>
        public decimal TI_HD27 { get; set; }
        /// <summary>
        /// TI諧波_28
        /// </summary>
        public decimal TI_HD28 { get; set; }
        /// <summary>
        /// TI諧波_29
        /// </summary>
        public decimal TI_HD29 { get; set; }
        /// <summary>
        /// TI諧波_30
        /// </summary>
        public decimal TI_HD30 { get; set; }
        /// <summary>
        /// TI諧波_31
        /// </summary>
        public decimal TI_HD31 { get; set; }
        #endregion
        /// <summary>
        /// 累積功率
        /// </summary>
        public decimal KWH { get; set; }
    }
}
