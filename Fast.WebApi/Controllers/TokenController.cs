using Fast.Application.Interfaces;
using Fast.Infra.CrossCutting.Common;
using Fast.WebApi.JWT;
using Fast.WebApi.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace Fast.WebApi.Controllers
{
    public class TokenController : ApiController
    {
        private readonly IUserAppService _userAppService;

        public TokenController(IUserAppService userAppService)
        {
            _userAppService = userAppService;
        }

        [AllowAnonymous]
        public HttpResponseMessage Post([FromBody] LoginModel loginModel)
        {
            try
            {
                if (ValidateUser(loginModel))
                {
                    string token = JwtManager.GenerateToken(loginModel.UserName);

                    JObject tokenResponse = new JObject(
                                                            new JProperty("username", loginModel.UserName),
                                                            new JProperty("token", token),
                                                            new JProperty("token_type", "bearer"));

                    return Request.CreateResponse(HttpStatusCode.OK, tokenResponse, Configuration.Formatters.JsonFormatter);
                }
                else
                {
                    HttpResponseMessage response = this.Request.CreateErrorResponse(HttpStatusCode.Unauthorized, "Username dan Password tidak valid");
                    throw new HttpResponseException(response);
                }
            }
            catch (Exception ex)
            {
                HttpResponseMessage response = this.Request.CreateErrorResponse(HttpStatusCode.InternalServerError, ex.Message);
                throw new HttpResponseException(response);
            }
        }

        private bool ValidateUser(LoginModel loginModel)
        {
            ICollection<QueryFilter> filters = new List<QueryFilter>();
            filters.Add(new QueryFilter("UserName", loginModel.UserName, Operator.Equals));

            string user = _userAppService.Get(filters);
            if (string.IsNullOrEmpty(user))
                return false;

            UserModel userModel = JsonConvert.DeserializeObject<UserModel>(user);
            if (loginModel.Password != AuthHelper.Decrypt(userModel.Password))
                return false;

            return true;
        }
    }
}
