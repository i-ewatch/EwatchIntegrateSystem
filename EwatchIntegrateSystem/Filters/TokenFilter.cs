using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using Microsoft.Extensions.Primitives;

namespace EwatchIntegrateSystem.Filters
{
    /// <summary>
    /// Token篩選功能
    /// </summary>
    public class TokenFilter : Attribute, IAuthorizationFilter
    {
        /// <summary>
        /// 啟用內容
        /// </summary>
        /// <param name="context"></param>
        public void OnAuthorization(AuthorizationFilterContext context)
        {
            bool TokenFlag = context.HttpContext.Request.Headers.TryGetValue("Authorization", out StringValues Token);
            var Ignore = (from a in context.ActionDescriptor.EndpointMetadata where a.GetType() == typeof(AllowAnonymousAttribute) select a).FirstOrDefault();
            if (Ignore == null)
            {
                if (TokenFlag)
                {
                    var tDll = GetType().Assembly.GetName();
                    if (tDll.Version != null)
                    {
                        if (Token.ToString() != tDll.Version.ToString())
                        {
                            context.Result = new JsonResult(new ResultJson()
                            {
                                Data = "Error",
                                HttpCode = 401,
                                Error = "版本錯誤!"
                            });
                        }
                    }
                }
                else
                {
                    context.Result = new JsonResult(new ResultJson()
                    {
                        Data = "Error",
                        HttpCode = 401,
                        Error = "請登入帳號!"
                    });
                }
            }
        }
    }
}
