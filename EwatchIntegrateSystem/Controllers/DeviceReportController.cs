using EwatchIntegrateSystem.Enums;
using EwatchIntegrateSystem.Methods;
using EwatchIntegrateSystem.Modules;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace EwatchIntegrateSystem.Controllers
{
    /// <summary>
    /// 設備報表分析
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class DeviceReportController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private ElectricLogMethod ElectricLogMethod { get; set; }
        /// <summary>
        /// 設備報表分析
        /// </summary>
        /// <param name="configuration"></param>
        public DeviceReportController(IConfiguration configuration)
        {
            _configuration = configuration;
            ElectricLogMethod = new ElectricLogMethod(_configuration["SqlDB"]);
        }
        /// <summary>
        /// 項目
        /// </summary>
        /// <param name="deviceTypeEnum"></param>
        /// <returns></returns>
        [Route("Kind")]
        [HttpGet]
        public IActionResult Kind(int deviceTypeEnum)
        {
            DeviceTypeEnum device = (DeviceTypeEnum)deviceTypeEnum;
            switch (device)
            {
                case DeviceTypeEnum.HC6600:
                case DeviceTypeEnum.PA310:
                case DeviceTypeEnum.PA60:
                case DeviceTypeEnum.TOYOPM96:
                case DeviceTypeEnum.PM5570:
                    {
                        return Ok(new
                        {
                            data = new string[] {
                                "三相電壓",
                                "三線電壓",
                                "三相電流",
                                "頻率",
                                "即時需量",
                                "虛功率",
                                "視在功率",
                                "功率因數",
                                "累積功率"
                            }
                        });
                    }
                case DeviceTypeEnum.PA3000:
                    {
                        return Ok(new
                        {
                            data = new string[] {
                                "三相電壓",
                                "三線電壓",
                                "三相電流",
                                "頻率",
                                "即時需量",
                                "虛功率",
                                "視在功率",
                                "功率因數",
                                "電壓相位角",
                                "電流相位角",
                                "R相電壓諧波",
                                "S相電壓諧波",
                                "T相電壓諧波",
                                "R相電流諧波",
                                "S相電流諧波",
                                "T相電流諧波",
                                "累積功率"
                            }
                        });
                    }
                default:
                    {
                        return BadRequest("查無該設備查詢種類");
                    }
            }
        }
        /// <summary>
        /// 項次
        /// </summary>
        /// <param name="Kind">種類</param>
        /// <returns></returns>
        [Route("Item")]
        [HttpGet]
        public IActionResult Item(string Kind)
        {
            switch (Kind)
            {
                case "三相電壓":
                case "三相電流":
                case "即時虛量":
                case "虛功率":
                case "視在功率":
                case "功率因數":
                case "電壓相位角":
                case "電流相位角":
                    {
                        return Ok(new
                        {
                            Name = new string[] { "全部", "R", "S", "T", },
                            data = new int[] { 0, 1, 2, 3 }
                        });
                    }
                case "三線電壓":
                    {
                        return Ok(new
                        {
                            Name = new string[] { "全部", "RS", "ST", "TR", "平均" },
                            data = new int[] { 0, 1, 2, 3, 4 }
                        });
                    }
                case "頻率":
                case "累積功率":
                    {
                        return Ok(new
                        {
                            Name = new string[] { "無" },
                            data = new int[] { 0 }
                        });
                    }
                case "R相電壓諧波":
                case "S相電壓諧波":
                case "T相電壓諧波":
                case "R相電流諧波":
                case "S相電流諧波":
                case "T相電流諧波":
                    {
                        return Ok(new
                        {
                            Name = new string[] { "全部",
                                "HD1",
                                "HD2",
                                "HD3",
                                "HD4",
                                "HD5",
                                "HD6",
                                "HD7",
                                "HD8",
                                "HD9",
                                "HD10",
                                "HD11",
                                "HD12",
                                "HD13",
                                "HD14",
                                "HD15",
                                "HD16",
                                "HD17",
                                "HD18",
                                "HD19",
                                "HD20",
                                "HD21",
                                "HD22",
                                "HD23",
                                "HD24",
                                "HD25",
                                "HD26",
                                "HD27",
                                "HD28",
                                "HD29",
                                "HD30",
                                "HD31"
                            },
                            data = new int[] {
                                0,
                                1,
                                2,
                                3,
                                4,
                                5,
                                6,
                                7,
                                8,
                                9,
                                10,
                                11,
                                12,
                                13,
                                14,
                                15,
                                16,
                                17,
                                18,
                                19,
                                20,
                                21,
                                22,
                                23,
                                24,
                                25,
                                26,
                                27,
                                28,
                                29,
                                30,
                                31,
                            }
                        });
                    }
                default:
                    {
                        return BadRequest("查無該設備項次");
                    }
            }
        }
        /// <summary>
        /// 查詢設備曲線圖
        /// </summary>
        /// <param name="StartTime">開始時間</param>
        /// <param name="EndTime">結束時間</param>
        /// <param name="deviceSetting">設備資訊</param>
        /// <param name="Kind">種類</param>
        /// <param name="Item">項次</param>
        /// <returns></returns>
        [Route("Search_Data")]
        [HttpPost]
        public IActionResult Search_Data(DateTime StartTime, DateTime EndTime, DeviceSetting deviceSetting, string Kind, int Item)
        {
            if (StartTime > EndTime)
                return BadRequest("輸入錯誤的查詢區間條件");
            var data = ElectricLogMethod.Search_Chart(StartTime, EndTime, deviceSetting, Kind, Item);
            return Ok(data);
        }
        /// <summary>
        /// 下載設備報表
        /// </summary>
        /// <param name="StartTime">開始時間</param>
        /// <param name="EndTime">結束時間</param>
        /// <param name="deviceSetting">設備資訊</param>
        /// <param name="Kind">種類</param>
        /// <param name="Item">項次</param>
        /// <returns></returns>
        [Route("DownLoad_Data")]
        [HttpPost]
        public IActionResult DownLoad_Data(DateTime StartTime, DateTime EndTime, DeviceSetting deviceSetting, string Kind, int Item)
        {
            if (StartTime > EndTime)
                return BadRequest("輸入錯誤的查詢區間條件");
            try
            {
                FileStreamResult? result = ElectricLogMethod.DownLoad_Device_Chart_Report(StartTime, EndTime, deviceSetting, Kind, Item);
                if (result != null)
                {
                    return result;
                }
                else
                {
                    return BadRequest("查詢報表失敗");
                }
            }
            catch (Exception)
            {
                return BadRequest("查詢報表失敗");
            }
        }
    }
}
