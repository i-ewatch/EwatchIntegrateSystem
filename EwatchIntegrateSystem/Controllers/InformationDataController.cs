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
    /// 數值資訊
    /// </summary>
    [Route("api/[controller]")]
    [ApiController]
    public class InformationDataController : ControllerBase
    {
        private readonly IHubContext<InformationDataHub> _hubContext;
        private readonly IConfiguration _configuration;
        private WebMethod WebMethod { get; set; }
        private ElectricLogMethod ElectricLogMethod { get; set; }
        /// <summary>
        /// 數值資訊初始化
        /// </summary>
        /// <param name="configuration"></param>
        /// <param name="hubContext"></param>
        public InformationDataController(IConfiguration configuration,IHubContext<InformationDataHub> hubContext)
        {
            _configuration = configuration;
            _hubContext = hubContext;
            WebMethod = new WebMethod(_configuration["SqlDB"]);
            ElectricLogMethod = new ElectricLogMethod(_configuration["SqlDB"]);
        }
        #region 回應訊息處理
        /// <summary>
        /// 回應訊息處理
        /// </summary>
        /// <param name="result">回應物件</param>
        /// <param name="Flag">訊息類別</param>
        /// <param name="Message">訊息字串</param>
        /// <param name="setting">取得物件</param>
        private void ResponseMessage(ResultJson result, ResponseTypeEnum Flag, string Message, object? setting)
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
                case ResponseTypeEnum.NoneCardBord:
                    {
                        result.Error = $"{Message}沒有建立卡版號";
                        result.HttpCode = 402;
                        result.Data = setting;
                    }
                    break;
            }
        }
        #endregion

        #region Web即時資訊
        /// <summary>
        /// 取得全部電表即時資訊
        /// </summary>
        /// <returns></returns>
        [Route("ElectricWeb")]
        [HttpGet]
        public async Task<IActionResult> GetAllElectricWeb()
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    List<ElectricWeb> cases = WebMethod.ElectricWeb_Search();
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
        /// 取得特定卡版號全部電表即時資訊
        /// </summary>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        [Route("ElectricWeb/{CardNumber}/{BordNumber}")]
        [HttpGet]
        public async Task<IActionResult> GetAllElectricWeb(string CardNumber, string BordNumber)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    List<ElectricWeb> cases = WebMethod.ElectricWeb_Search(CardNumber, BordNumber);
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
        /// 取得最新驟降電壓資訊
        /// </summary>
        /// <param name="StartTime">當月時間 (yyyyMM)</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        [Route("ElectricSagWeb/{StartTime}/{CardNumber}/{BordNumber}")]
        [HttpGet]
        public async Task<IActionResult> GetElectricSagWeb(string StartTime, string CardNumber, string BordNumber)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ElectricSag_Log? cases = ElectricLogMethod.ElectricSag_Search(StartTime, CardNumber, BordNumber);
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
        /// 取得最新驟升電壓資訊
        /// </summary>
        /// <param name="StartTime">當月時間 (yyyyMM)</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        [Route("ElectricSwellWeb/{StartTime}/{CardNumber}/{BordNumber}")]
        [HttpGet]
        public async Task<IActionResult> GetElectricSwellWeb(string StartTime, string CardNumber, string BordNumber)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ElectricSwell_Log? cases = ElectricLogMethod.ElectricSwell_Search(StartTime, CardNumber, BordNumber);
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
        /// 取得最新電表告警資訊
        /// </summary>
        /// <param name="StartTime">當月時間 (yyyyMM)</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        [Route("ElectricAlarmWeb/{StartTime}/{CardNumber}/{BordNumber}")]
        [HttpGet]
        public async Task<IActionResult> GetElectricAlarmWeb(string StartTime, string CardNumber, string BordNumber)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ElectricAlarm_Log? cases = ElectricLogMethod.ElectricAlarm_Search(StartTime, CardNumber, BordNumber);
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
        #endregion

        #region 歷史資訊
        /// <summary>
        /// 當月電表歷史資訊
        /// </summary>
        /// <param name="StartTime">當月時間 (yyyyMM)</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        [Route("ElectricInformationLog/{StartTime}/{CardNumber}/{BordNumber}")]
        [HttpGet]
        public async Task<IActionResult> GetElectricLog(string StartTime, string CardNumber, string BordNumber)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    List<ElectricInformation_Log> cases = ElectricLogMethod.ElectricInformation_SearchMonth(StartTime, CardNumber, BordNumber);
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
        /// 當月驟降電壓歷史資訊
        /// </summary>
        /// <param name="StartTime">當月時間 (yyyyMM)</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        [Route("ElectricSagLog/{StartTime}/{CardNumber}/{BordNumber}")]
        [HttpGet]
        public async Task<IActionResult> GetElectricSagLog(string StartTime, string CardNumber, string BordNumber)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    List<ElectricSag_Log> cases = ElectricLogMethod.ElectricSag_SearchMonth(StartTime, CardNumber, BordNumber);
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
        /// 當月驟升電壓歷史資訊
        /// </summary>
        /// <param name="StartTime">當月時間 (yyyyMM)</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        [Route("ElectricSwellLog/{StartTime}/{CardNumber}/{BordNumber}")]
        [HttpGet]
        public async Task<IActionResult> GetElectricSwellLog(string StartTime, string CardNumber, string BordNumber)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    List<ElectricSwell_Log> cases = ElectricLogMethod.ElectricSwell_SearchMonth(StartTime, CardNumber, BordNumber);
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
        /// 當月電表告警歷史資訊
        /// </summary>
        /// <param name="StartTime">當月時間 (yyyyMM)</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        [Route("ElectricAlarmLog/{StartTime}/{CardNumber}/{BordNumber}")]
        [HttpGet]
        public async Task<IActionResult> GetElectricAlarmLog(string StartTime, string CardNumber, string BordNumber)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    List<ElectricAlarm_Log> cases = ElectricLogMethod.ElectricAlarm_SearchMonth(StartTime, CardNumber, BordNumber);
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
        /// 查詢電表累積量資料(特定卡版號單顆電表)
        /// </summary>
        /// <param name="StartTime">開始時間(yyyyMMdd)</param>
        /// <param name="EndTime">結束時間(yyyyMMdd)</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <param name="DeviceNumber">設備編號</param>
        /// <returns></returns>
        [Route("ELectricTotal/{StartTime}/{EndTime}/{CardNumber}/{BordNumber}/{DeviceNumber}")]
        [HttpGet]
        public async Task<IActionResult> GetELectricTotal(string StartTime, string EndTime, string CardNumber, string BordNumber, int DeviceNumber)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    List<ElectricTotal_Log> cases = ElectricLogMethod.ElectricTotal_Search(StartTime, EndTime, CardNumber, BordNumber, DeviceNumber);
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
        #endregion

        #region ElectricLog資訊新增
        /// <summary>
        /// 新增電表資訊、更新電表即時資訊、刷新累積電量資訊
        /// </summary>
        /// <param name="electricInformation_Log">電表資訊</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        [Route("Electric")]
        [HttpPost]
        public async Task<IActionResult> InsertElectricInformationLog(ElectricInformation_Log electricInformation_Log, string CardNumber, string BordNumber)
        {

            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {
                    ResponseTypeEnum webFlag = ElectricLogMethod.ElectricWeb_AddUpdate(electricInformation_Log, CardNumber, BordNumber);//新增或更新電表即時資訊(Web)
                    ResponseTypeEnum Flag = ElectricLogMethod.ElectricInformation_Add(electricInformation_Log, CardNumber, BordNumber);//新增歷史數值
                    ResponseTypeEnum TotalFlag = ElectricLogMethod.ElectricTotal_Add(electricInformation_Log, CardNumber, BordNumber);//計算累積量
                    if (webFlag != ResponseTypeEnum.Finish && Flag != ResponseTypeEnum.Finish && TotalFlag != ResponseTypeEnum.Finish)
                    {
                        ResponseMessage(result, Flag, "新增電表、更新電表、累積電量數值", electricInformation_Log);
                    }
                    else if (webFlag == ResponseTypeEnum.Finish && Flag == ResponseTypeEnum.Finish && TotalFlag == ResponseTypeEnum.Finish)
                    {
                        ResponseMessage(result, Flag, "新增電表、更新電表、累積電量數值", electricInformation_Log);
                    }
                    else if (Flag != ResponseTypeEnum.Finish)
                    {
                        ResponseMessage(result, Flag, "新增電表數值", electricInformation_Log);
                    }
                    else if (webFlag != ResponseTypeEnum.Finish)
                    {
                        ResponseMessage(result, webFlag, "更新電表即時數值", electricInformation_Log);
                    }
                    else if (TotalFlag != ResponseTypeEnum.Finish)
                    {
                        ResponseMessage(result, TotalFlag, "刷新累積電量數值", electricInformation_Log);
                    }
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
        /// 新增電表驟降電壓資訊
        /// </summary>
        /// <param name="electricSag_Log">電表驟降電壓資訊</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        [Route("ElectricSag")]
        [HttpPost]
        public async Task<IActionResult> InserElectricSagLog(ElectricSag_Log electricSag_Log, string CardNumber, string BordNumber)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {

                    ResponseTypeEnum Flag = ElectricLogMethod.ElectricSag_Add(electricSag_Log, CardNumber, BordNumber);//新增歷史數值
                    ResponseMessage(result, Flag, "電表驟降電壓數值", electricSag_Log);
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
        /// 新增電表驟升電壓資訊
        /// </summary>
        /// <param name="electricSwell_Log">電表驟升電壓資訊</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        [Route("ElectricSwell")]
        [HttpPost]
        public async Task<IActionResult> InserElectricSwellLog(ElectricSwell_Log electricSwell_Log, string CardNumber, string BordNumber)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {

                    ResponseTypeEnum Flag = ElectricLogMethod.ElectricSwell_Add(electricSwell_Log, CardNumber, BordNumber);//新增歷史數值
                    ResponseMessage(result, Flag, "電表驟升電壓數值", electricSwell_Log);
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
        /// 新增電表告警資訊
        /// </summary>
        /// <param name="electricAlarm_Log">電表告警資訊</param>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        [Route("ElectricAlarm")]
        [HttpPost]
        public async Task<IActionResult> InserElectricAlarmLog(ElectricAlarm_Log electricAlarm_Log, string CardNumber, string BordNumber)
        {
            ResultJson result = new ResultJson();
            try
            {
                await Task.Run(() =>
                {

                    ResponseTypeEnum Flag = ElectricLogMethod.ElectricAlarm_Add(electricAlarm_Log, CardNumber, BordNumber);//新增歷史數值
                    ResponseMessage(result, Flag, "電表告警數值", electricAlarm_Log);
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
