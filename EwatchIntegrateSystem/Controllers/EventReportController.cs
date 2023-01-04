using EwatchIntegrateSystem.Enums;
using EwatchIntegrateSystem.Methods;
using EwatchIntegrateSystem.Modules;
using Microsoft.AspNetCore.Mvc;

namespace EwatchIntegrateSystem.Controllers
{
    /// <summary>
    /// 事件報表分析
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class EventReportController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private ElectricLogMethod ElectricLogMethod { get; set; }
        /// <summary>
        /// 事件報表分析
        /// </summary>
        /// <param name="configuration"></param>
        public EventReportController(IConfiguration configuration)
        {
            _configuration = configuration;
            ElectricLogMethod = new ElectricLogMethod(_configuration["SqlDB"]);
        }
        /// <summary>
        /// 項目
        /// </summary>
        /// <param name="deviceTypeEnum">設備類型</param>
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
                        return Ok("無");
                    }
                case DeviceTypeEnum.PA3000:
                    {
                        return Ok(new
                        {
                            data = new string[] {
                                "驟降驟升事件",
                                "需量&告警紀錄"
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
        /// 查詢設備事件
        /// </summary>
        /// <param name="StartTime">開始時間</param>
        /// <param name="EndTime">結束時間</param>
        /// <param name="deviceSetting">設備資訊</param>
        /// <param name="Kind">種類</param>
        /// <param name="Item">項次</param>
        /// <returns></returns>
        [Route("Search_Data")]
        [HttpPost]
        public IActionResult Search_Data(DateTime StartTime, DateTime EndTime, DeviceSetting deviceSetting, string Kind)
        {
            if (StartTime > EndTime)
                return BadRequest("輸入錯誤的查詢區間條件");
            var data = ElectricLogMethod.Search_Event(StartTime, EndTime, deviceSetting, Kind);
            return Ok(data);
        }
        /// <summary>
        /// 下載設備事件報表
        /// </summary>
        /// <param name="StartTime">開始時間</param>
        /// <param name="EndTime">結束時間</param>
        /// <param name="deviceSetting">設備資訊</param>
        /// <param name="Kind">種類</param>
        /// <returns></returns>
        [Route("DownLoad_Data")]
        [HttpPost]
        public IActionResult DownLoad_Data(DateTime StartTime, DateTime EndTime, DeviceSetting deviceSetting, string Kind)
        {
            if (StartTime > EndTime)
                return BadRequest("輸入錯誤的查詢區間條件");
            try
            {
                FileStreamResult? result = ElectricLogMethod.DownLoad_Device_Event_Report(StartTime, EndTime, deviceSetting, Kind);
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
