using EwatchIntegrateSystem.Methods;
using EwatchIntegrateSystem.Modules;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Metadata;
using System.Threading.Tasks;

namespace EwatchIntegrateSystem.Controllers
{
    /// <summary>
    /// 電表資訊
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class ElectricController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private WebMethod WebMethod { get; set; }
        private ElectricLogMethod ElectricLogMethod { get; set; }
        /// <summary>
        /// 電表資訊
        /// </summary>
        /// <param name="configuration"></param>
        public ElectricController(IConfiguration configuration)
        {
            _configuration = configuration;
            WebMethod = new WebMethod(_configuration["SqlDB"]);
            ElectricLogMethod = new ElectricLogMethod(_configuration["SqlDB"]);
        }
        /// <summary>
        /// 月用電趨勢比較
        /// </summary>
        /// <param name="search_Electric">查詢電表資訊</param>
        /// <returns></returns>
        [Route("Device_Month_kWh")]
        [HttpPost]
        public IActionResult Device_Month_kWh(Search_Electric search_Electric)
        {
            var data = ElectricLogMethod.Search_Month_kWh(search_Electric.StartTime, search_Electric.CardNumber, search_Electric.BordNumber, search_Electric.DeviceName, search_Electric.DeviceNumber);
            IList<kWh_Bar> value = data.ToList<kWh_Bar>();
            List<decimal> NowValue = new List<decimal>();
            List<decimal> LastValue = new List<decimal>();
            List<string> XAxis = new List<string>();
            List<object> Series = new List<object>();
            XAxis.AddRange(value.Select(t => t.ttime).ToList());
            Series.Add(new
            {
                Name = "今年月累積量",
                data = value.Select(v => v.NowTotalkWh).ToList(),
                Type = "bar"
            });
            Series.Add(new
            {
                Name = "去年月累積量",
                data = value.Select(v => v.AfterTotalkWh).ToList(),
                Type = "bar"
            });
            return Ok(new
            {
                XAxis = XAxis,
                Series = Series
            });
        }
        /// <summary>
        /// 年用電趨勢比較
        /// </summary>
        /// <param name="search_Electric">查詢電表資訊</param>
        /// <returns></returns>
        [Route("Device_Year_kWh")]
        [HttpPost]
        public IActionResult Device_Year_kWh(Search_Electric search_Electric)
        {
            var data = ElectricLogMethod.Search_Year_kWh(search_Electric.StartTime, search_Electric.CardNumber, search_Electric.BordNumber, search_Electric.DeviceName, search_Electric.DeviceNumber);
            IList<kWh_Bar> value = data.ToList<kWh_Bar>();
            List<decimal> NowValue = new List<decimal>();
            List<decimal> LastValue = new List<decimal>();
            List<string> XAxis = new List<string>();
            List<object> Series = new List<object>();
            XAxis.AddRange(value.Select(t => t.ttime).ToList());
            Series.Add(new
            {
                Name = "今年累積量",
                data = value.Select(v => v.NowTotalkWh).ToList(),
                Type = "bar"
            });
            Series.Add(new
            {
                Name = "去年累積量",
                data = value.Select(v => v.AfterTotalkWh).ToList(),
                Type = "bar"
            });
            return Ok(new
            {
                XAxis = XAxis,
                Series = Series
            });
        }
        private class kWh_Bar
        {
            /// <summary>
            /// 時間
            /// </summary>
            public string ttime { get; set; } = string.Empty;
            /// <summary>
            /// 今年本月累積量
            /// </summary>
            public string NowTotalkWh { get; set; } = string.Empty;
            /// <summary>
            /// 去年本月累積量
            /// </summary>
            public string AfterTotalkWh { get; set; } = string.Empty;
        }
        /// <summary>
        /// 用電累積量資訊
        /// </summary>
        /// <param name="search_Electric">查詢電表資訊</param>
        /// <returns></returns>
        [Route("Device_Total_Info")]
        [HttpPost]
        public IActionResult Device_Total_Info(Search_Electric search_Electric)
        {
            var data = ElectricLogMethod.Search_Total_Info(search_Electric.StartTime, search_Electric.CardNumber, search_Electric.BordNumber, search_Electric.DeviceName, search_Electric.DeviceNumber);
            return Ok(data);
        }
        /// <summary>
        /// 驟降驟升事件
        /// </summary>
        /// <param name="search_Electric">查詢電表資訊</param>
        /// <returns></returns>
        [Route("Device_Sag_Swell_Info")]
        [HttpPost]
        public IActionResult Device_Sag_Swell_Info(Search_Electric search_Electric)
        {
            var data = ElectricLogMethod.Search_Sag_Swell_Info(search_Electric.CardNumber, search_Electric.BordNumber, search_Electric.DeviceNumber);
            return Ok(data);
        }
        /// <summary>
        /// 超約告警事件
        /// </summary>
        /// <param name="search_Electric">查詢電表資訊</param>
        /// <returns></returns>
        [Route("Device_Alarm_Info")]
        [HttpPost]
        public IActionResult Device_Alarm_Info(Search_Electric search_Electric)
        {
            var data = ElectricLogMethod.Search_Alarm_Info(search_Electric.CardNumber, search_Electric.BordNumber, search_Electric.DeviceNumber);
            return Ok(data);
        }
    }
}
