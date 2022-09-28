using EwatchIntegrateSystem.Modules;
using Microsoft.AspNetCore.SignalR;

namespace EwatchIntegrateSystem.Hubs
{
    /// <summary>
    /// Signalr 設備資訊Hub
    /// </summary>
    public class InformationDataHub:Hub
    {
        /// <summary>
        /// 呼叫特定卡版號全部電表即時資訊(目前不使用)
        /// </summary>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        public async Task Electric_WhoIs(string CardNumber, string BordNumber)
        {
            await Clients.All.SendAsync($"Electric_WhoIs", CardNumber,BordNumber);
        }
        /// <summary>
        /// 現場特定卡版號全部電表即時資訊
        /// </summary>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <param name="electricInformation_Logs">電表資訊</param>
        /// <returns></returns>
        public async Task ElectricInformationData(string CardNumber, string BordNumber,List<ElectricInformation_Log> electricInformation_Logs)
        {
            await Clients.All.SendAsync($"Electric_{CardNumber}{BordNumber}", electricInformation_Logs);
        }
    }
}
