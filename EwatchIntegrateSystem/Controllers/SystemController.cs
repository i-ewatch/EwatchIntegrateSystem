using EwatchIntegrateSystem.Enums;
using EwatchIntegrateSystem.Filters;
using EwatchIntegrateSystem.Hubs;
using EwatchIntegrateSystem.Methods;
using EwatchIntegrateSystem.Modules;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.SignalR;

namespace EwatchIntegrateSystem.Controllers
{
    /// <summary>
    /// 基本資訊
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class SystemController : ControllerBase
    {
        private readonly IHubContext<SystemHub> _hubContext;
        private readonly IConfiguration _configuration;
        /// <summary>
        /// 案件資料庫方法
        /// </summary>
        private CaseMethod CaseMethod { get; set; }
        /// <summary>
        /// 設備資料庫方法
        /// </summary>
        private DeviceMethod DeviceMethod { get; set; }
        /// <summary>
        /// 設備類型資料庫方法
        /// </summary>
        private DeviceTypeEnumMethod DeviceTypeEnumMethod { get; set; }
        /// <summary>
        /// 基本資訊初始化
        /// </summary>
        public SystemController(IConfiguration configuration, IHubContext<SystemHub> hubContext)
        {
            _configuration = configuration;
            _hubContext = hubContext;
            CaseMethod = new CaseMethod(_configuration["SqlDB"]);
            DeviceMethod = new DeviceMethod(_configuration["SqlDB"]);
            DeviceTypeEnumMethod = new DeviceTypeEnumMethod(_configuration["SqlDB"]);
        }
        #region 回應訊息處理
        /// <summary>
        /// 回應訊息處理
        /// </summary>
        /// <param name="result">回應物件</param>
        /// <param name="Flag">訊息類別</param>
        /// <param name="Message">訊息字串</param>
        /// <param name="setting">取得物件</param>
        private void ResponseMessage(ResultJson result, ResponseTypeEnum Flag, string Message, object setting)
        {
            switch (Flag)
            {
                case ResponseTypeEnum.Repeat:
                    {
                        result.Error = $"{Message}重複";
                        result.HttpCode = 401;
                    }
                    break;
                case ResponseTypeEnum.Finish:
                    {
                        result.Error = $"{Message}完成";
                        result.HttpCode = 200;
                        result.Data = setting;
                    }
                    break;
                case ResponseTypeEnum.Error:
                    {
                        result.Error = $"{Message}失敗";
                        result.HttpCode = 401;
                    }
                    break;
            }
        }
        #endregion
        #region 案件控制物件
        /// <summary>
        /// 取得所有案件資訊
        /// </summary>
        /// <returns></returns>
        [Route("Case")]
        [HttpGet]
        public async Task<IActionResult> GetAllCasse()
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    List<CaseSetting> cases = CaseMethod.Case_Search();
                    ResponseMessage(result, ResponseTypeEnum.Finish, "", cases);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        /// <summary>
        /// 新增案件資訊
        /// </summary>
        /// <param name="CaseSetting">案件資訊</param>
        /// <returns></returns>
        [Route("Case")]
        [HttpPost]
        public async Task<IActionResult> InserterCase(CaseSetting CaseSetting)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ResponseTypeEnum Flag = CaseMethod.Case_Add(CaseSetting);
                    ResponseMessage(result, Flag, "卡版號新增", CaseSetting);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        /// <summary>
        /// 修改案件資訊
        /// </summary>
        /// <param name="CaseSetting">案件資訊</param>
        /// <returns></returns>
        [Route("Case")]
        [HttpPut]
        public async Task<IActionResult> UpdateCase(CaseSetting CaseSetting)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ResponseTypeEnum Flag = CaseMethod.Case_Update(CaseSetting);
                    ResponseMessage(result, Flag, "卡版號修改", CaseSetting);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        /// <summary>
        /// 刪除案件資訊
        /// </summary>
        /// <param name="CaseSetting">案件資訊</param>
        /// <returns></returns>
        [Route("Case")]
        [HttpDelete]
        public async Task<IActionResult> DeleteCase(CaseSetting CaseSetting)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ResponseTypeEnum Flag = CaseMethod.Case_Delete(CaseSetting);
                    ResponseMessage(result, Flag, "卡版號刪除", CaseSetting);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        #endregion
        #region 設備控制物件
        /// <summary>
        /// 取得所有設備資訊
        /// </summary>
        /// <returns></returns>
        [Route("Device")]
        [HttpGet]
        public async Task<IActionResult> GetAllDevice()
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    List<DeviceSetting> settings = DeviceMethod.Device_Search();
                    ResponseMessage(result, ResponseTypeEnum.Finish, "", settings);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        /// <summary>
        /// 取得特定卡版號設備資訊
        /// </summary>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        [Route("Device/{CardNumber}/{BordNumber}")]
        [HttpGet]
        public async Task<IActionResult> GetAllDevice(string CardNumber, string BordNumber)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    List<DeviceSetting> settings = DeviceMethod.Device_Search(CardNumber, BordNumber);
                    ResponseMessage(result, ResponseTypeEnum.Finish, "", settings);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        /// <summary>
        /// 新增單個設備資訊
        /// </summary>
        /// <param name="deviceSetting">設備資訊</param>
        /// <returns></returns>
        [Route("Device/Single")]
        [HttpPost]
        public async Task<IActionResult> InsertDevice(DeviceSetting deviceSetting)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ResponseTypeEnum Flag = DeviceMethod.Device_Single_Add(deviceSetting);
                    ResponseMessage(result, Flag, "設備單個新增", deviceSetting);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        /// <summary>
        /// 新增多個設備資訊
        /// </summary>
        /// <param name="deviceSetting">設備資訊</param>
        /// <returns></returns>
        [Route("Device/Multiple")]
        [HttpPost]
        public async Task<IActionResult> InsertDevice(List<DeviceSetting> deviceSetting)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ResponseTypeEnum Flag = DeviceMethod.Device_Multiple_Add(deviceSetting);
                    ResponseMessage(result, Flag, "設備多個新增", deviceSetting);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        /// <summary>
        /// 修改單個設備資訊
        /// </summary>
        /// <param name="deviceSetting">設備資訊</param>
        /// <returns></returns>
        [Route("Device/Single")]
        [HttpPut]
        public async Task<IActionResult> UpdateDevice(DeviceSetting deviceSetting)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ResponseTypeEnum Flag = DeviceMethod.Device_Single_Update(deviceSetting);
                    ResponseMessage(result, Flag, "設備單個修改", deviceSetting);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        /// <summary>
        /// 修改多個設備資訊
        /// </summary>
        /// <param name="deviceSetting">設備資訊</param>
        /// <returns></returns>
        [Route("Device/Multiple")]
        [HttpPut]
        public async Task<IActionResult> UpdateDevice(List<DeviceSetting> deviceSetting)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ResponseTypeEnum Flag = DeviceMethod.Device_Multiple_Update(deviceSetting);
                    ResponseMessage(result, Flag, "設備多個修改", deviceSetting);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        /// <summary>
        /// 刪除單個設備資訊
        /// </summary>
        /// <param name="deviceSetting">設備資訊</param>
        /// <returns></returns>
        [Route("Device/Single")]
        [HttpDelete]
        public async Task<IActionResult> DeleteDevice(DeviceSetting deviceSetting)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ResponseTypeEnum Flag = DeviceMethod.Device_Single_Delete(deviceSetting);
                    ResponseMessage(result, Flag, "設備單個刪除", deviceSetting);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        /// <summary>
        /// 刪除多個設備資訊
        /// </summary>
        /// <param name="deviceSetting">設備資訊</param>
        /// <returns></returns>
        [Route("Device/Multiple")]
        [HttpDelete]
        public async Task<IActionResult> DeleteDevice(List<DeviceSetting> deviceSetting)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ResponseTypeEnum Flag = DeviceMethod.Device_Multiple_Delete(deviceSetting);
                    ResponseMessage(result, Flag, "設備多個刪除", deviceSetting);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        #endregion
        #region 設備類型控制物件
        /// <summary>
        /// 取得所有設備類型資訊
        /// </summary>
        /// <returns></returns>
        [Route("DeviceType")]
        [HttpGet]
        public async Task<IActionResult> GetAllDeviceType()
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    List<DeviceTypeEnumSetting> settings = DeviceTypeEnumMethod.DeviceType_Search();
                    ResponseMessage(result, ResponseTypeEnum.Finish, "", settings);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        /// <summary>
        /// 新增設備類型資訊
        /// </summary>
        /// <param name="deviceTypeEnumSetting">設備類型資訊</param>
        /// <returns></returns>
        [Route("DeviceType")]
        [HttpPost]
        public async Task<IActionResult> InsertDevice(DeviceTypeEnumSetting deviceTypeEnumSetting)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ResponseTypeEnum Falg = DeviceTypeEnumMethod.DeviceType_Add(deviceTypeEnumSetting);
                    ResponseMessage(result, Falg, "設備類型新增", deviceTypeEnumSetting);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        /// <summary>
        /// 修改設備類型資訊
        /// </summary>
        /// <param name="deviceTypeEnumSetting">設備類型資訊</param>
        /// <returns></returns>
        [Route("DeviceType")]
        [HttpPut]
        public async Task<IActionResult> UpdateDevice(DeviceTypeEnumSetting deviceTypeEnumSetting)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ResponseTypeEnum Falg = DeviceTypeEnumMethod.DeviceType_Update(deviceTypeEnumSetting);
                    ResponseMessage(result, Falg, "設備類型修改", deviceTypeEnumSetting);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        /// <summary>
        /// 刪除設備類型資訊
        /// </summary>
        /// <param name="deviceTypeEnumSetting">設備類型資訊</param>
        /// <returns></returns>
        [Route("DeviceType")]
        [HttpDelete]
        public async Task<IActionResult> DeleteDevice(DeviceTypeEnumSetting deviceTypeEnumSetting)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ResponseTypeEnum Falg = DeviceTypeEnumMethod.DeviceType_Delete(deviceTypeEnumSetting);
                    ResponseMessage(result, Falg, "設備類型刪除", deviceTypeEnumSetting);
                });
                return new JsonResult(result);
            }
            catch (Exception ex)
            {
                result.Error = ex.Message.ToString();
                return new JsonResult(result);
            }
        }
        #endregion
    }
}