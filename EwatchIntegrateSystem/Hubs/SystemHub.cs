using Microsoft.AspNetCore.SignalR;

namespace EwatchIntegrateSystem.Hubs
{
    /// <summary>
    /// Signalr 系統資訊Hub
    /// </summary>
    public class SystemHub:Hub
    {
        /// <summary>
        /// 呼叫指定卡版號資訊
        /// </summary>
        /// <param name="CardNumber">卡號</param>
        /// <param name="BordNumber">版號</param>
        /// <returns></returns>
        public async Task WhoIs(string CardNumber, string BordNumber)
        {
            await Clients.All.SendAsync("WhoIs", CardNumber+BordNumber);
        }
        /// <summary>
        /// 回應指定卡版號資訊
        /// </summary>
        /// <param name="CardNumber"></param>
        /// <param name="BordNumber"></param>
        /// <returns></returns>
        public async Task Iam(string CardNumber, string BordNumber)
        {
            await Clients.All.SendAsync("Iam", CardNumber+BordNumber);
        }
    }
}
