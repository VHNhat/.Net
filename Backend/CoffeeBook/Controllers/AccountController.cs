using CoffeeBook.Authen;
using CoffeeBook.DataAccess;
using CoffeeBook.Dto;
using CoffeeBook.Models;
using CoffeeBook.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Tokens;
using System;
using System.Collections.Generic;
using System.Data;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;

namespace CoffeeBook.Controllers
{
    /*[Route("api/[controller]")]*/
    [ApiController]
    public class AccountController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly AccountService _service;
        private readonly Context context;
        private readonly AppSetting _appSettings;

        public AccountController(IConfiguration config, Context ctx, IOptions<AppSetting> appSettings)
        {
            _config = config;
            context = ctx;
            _appSettings = appSettings.Value;
            _service = new AccountService(_config, ctx);
        }


        [Route("admin/login")]
        [HttpPost]
        public ActionResult Login(AdminLoginDto dto)
        {
            Account result = _service.Login(dto);

            if (result == null)
            {
                return new JsonResult("Username or Password is invalid.");
            }
            if (!BCrypt.Net.BCrypt.Verify(dto.Password, result.Password))
            {
                return BadRequest(new { message = "Invalid Credentials" });
            }

            var token = generateJwtToken(result);

            Response.Cookies.Append("jwt", token, new CookieOptions
            {
                HttpOnly = true
            });
            return new JsonResult(new { Token = token });
        }


        [Route("account")]
        [HttpGet]
        public ActionResult GetAllAccount()
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var role = getCurrentRole(jwt);
                if (role == "1")
                {
                    List<Account> accounts = _service.FindAll();
                    if (accounts.Count == 0)
                        return new JsonResult("There is no data");

                    return new JsonResult(accounts);
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập"});
        }

        [Route("account/{id}")]
        [HttpGet]
        public ActionResult GetAccountById(int id)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var role = getCurrentRole(jwt);
                if (!string.IsNullOrEmpty(role))
                {
                    Account account = _service.FindById(id);
                    if (account == null)
                        return BadRequest();

                    return new JsonResult(account);
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("account/add")]
        [HttpPost]
        public ActionResult AddAccount(Account account)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var role = getCurrentRole(jwt);
                if (role == "1")
                {
                    var newAcc = new Account 
                    {
                        Username = account.Username,
                        Password = BCrypt.Net.BCrypt.HashPassword(account.Password),
                        RoleId = account.RoleId,
                    };
                    int result = _service.Add(newAcc);
                    if (result > 0)
                        return Ok();
                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }


        [Route("admin/logout")]
        [HttpPost]
        public ActionResult Logout()
        {
            Response.Cookies.Delete("jwt");
            return Ok(new { message = "Đã đăng xuất" });
        }

        [Route("account/edit/{id}")]
        [HttpPut]
        public ActionResult UpdateAccount(int id, Account account)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var role = getCurrentRole(jwt);
                if (role == "1")
                {
                    if (ModelState.IsValid)
                    {
                        int res = _service.Update(id, account);
                        if (res > 0) return Ok();
                        else return BadRequest();
                    }
                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("account/delete/{id}")]
        [HttpDelete]
        public ActionResult DeleteAccount(int id)
        {

            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var role = getCurrentRole(jwt);
                if (role == "1")
                {
                    int res = _service.DeleteById(id);
                    if (res > 0) return Ok();
                    else return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }


        private string generateJwtToken(Account account)
        {
            var claims = new Claim[]
            {
                new Claim("Id", account.Id.ToString()),
                new Claim("Username", account.Username),
                new Claim("RoleId", account.RoleId.ToString()),
                new Claim(ClaimTypes.Role, account.RoleId.ToString())
            };

            // generate token that is valid for 7 days
            var tokenHandler = new JwtSecurityTokenHandler();
            var key = Encoding.ASCII.GetBytes(_appSettings.Secret);
            var tokenDescriptor = new SecurityTokenDescriptor
            {
                Subject = new ClaimsIdentity(claims),
                Expires = DateTime.UtcNow.AddMinutes(20),
                SigningCredentials = new SigningCredentials(new SymmetricSecurityKey(key), SecurityAlgorithms.HmacSha256Signature)
            };
            var token = tokenHandler.CreateToken(tokenDescriptor);
            return tokenHandler.WriteToken(token);
        }

        private string getCurrentRole(string jwt)
        {
            var handler = new JwtSecurityTokenHandler();
            var jsonToken = handler.ReadToken(jwt);
            var tokenS = jsonToken as JwtSecurityToken;
            var role = tokenS.Claims.First(claim => claim.Type == "RoleId").Value;
            return role;
        }
    }
}
