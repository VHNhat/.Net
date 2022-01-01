using CoffeeBook.Authen;
using CoffeeBook.DataAccess;
using CoffeeBook.Dto;
using CoffeeBook.Models;
using CoffeeBook.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.CodeAnalysis.Options;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Tokens;
using MySql.Data.MySqlClient;
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
    public class CustomerController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly CustomerService service;
        private readonly Context context;
        private readonly AppSetting _appSettings;

        public CustomerController(IConfiguration config, Context ctx, IOptions<AppSetting> appSettings)
        {
            _config = config;
            context = ctx;
            _appSettings = appSettings.Value;
            service = new CustomerService(_config, ctx);
        }


        [Route("customer")]
        [HttpGet]
        public ActionResult Get()
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1")
                {
                    List<Customer> customers = service.findAll();
                    return new JsonResult(customers);
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("customer/{id}")]
        [HttpGet]
        public ActionResult GetById(int id)
        {
            string jwt = Request.Cookies["jwt"];
            string jwtCustomer = Request.Cookies["jwtCustomer"];
            if (!string.IsNullOrEmpty(jwt) || !string.IsNullOrEmpty(jwtCustomer))
            {
                string Role = "", cusRole = "", cusId = "";
                if (!string.IsNullOrEmpty(jwt))
                {
                    Role = getCurrentRole(jwt);
                }
                if (!string.IsNullOrEmpty(jwtCustomer))
                {
                    cusRole = getCustomerRole(jwtCustomer);
                    cusId = getCustomerId(jwtCustomer);
                }
                if (Role == "1" || (cusRole == "Customer" && cusId == id.ToString()))
                {
                    Customer list = service.findById(id);
                    if (list == null) return BadRequest();
                    return new JsonResult(list);
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("customer/add")]
        [HttpPost]
        public ActionResult Post(Customer customer)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1")
                {
                    int res = service.save(customer);
                    if (res > 0) return Ok();

                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("customer/login")]
        [HttpPost]
        public ActionResult Login(SigninDto dto)
        {
            Customer user = service.Login(dto);

            if(user == null)
                return BadRequest(new { message = "Invalid Credentials" });

            if (!BCrypt.Net.BCrypt.Verify(dto.Password, user.Password))
            {
               
                return BadRequest(new { message = "Invalid Credentials"});
            }

            var token = generateJwtToken(user);

            Response.Cookies.Append("jwtCustomer", token, new CookieOptions
            {
                HttpOnly = true
            });

            return new JsonResult(new { Token = token });
        }

        [Route("customer/signup")]
        [HttpPost]
        public ActionResult Register(SignupDto dto)
        {
            string res = service.Register(dto);
            if (res == "1") return Ok();
            else return new JsonResult(res);
        }

        [Route("customer/logout")]
        [HttpPost]
        public ActionResult Logout()
        {
            Response.Cookies.Delete("jwtCustomer");
            return Ok(new { message = "Đã đăng xuất" });
        }

        [Route("customer/edit/{id}")]
        [HttpPut]
        public ActionResult Put(int id,Customer customer)
        {
            string jwt = Request.Cookies["jwt"];
            string jwtCustomer = Request.Cookies["jwtCustomer"];
            if (!string.IsNullOrEmpty(jwt) || !string.IsNullOrEmpty(jwtCustomer))
            {
                string Role = "", cusRole = "";
                if (!string.IsNullOrEmpty(jwt))
                {
                    Role = getCurrentRole(jwt);
                }
                if (!string.IsNullOrEmpty(jwtCustomer))
                {
                    cusRole = getCustomerRole(jwtCustomer);
                }
                
                if (Role == "1" || cusRole == "Customer")
                {
                    int res = service.update(id, customer);
                    if (res > 0) return Ok();
                    else return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }
        [Route("customer/forgot/{email}")]
        [HttpPut]
        public ActionResult ForgotPassword(string email, ForgotPassDto dto)
        {
            var res = service.ChangePassword(email, dto);
            if (res > 0)
                return Ok();
            return BadRequest();
        }

        [Route("customer/delete/{id}")]
        [HttpDelete]
        public ActionResult Delete(int id)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1")
                {
                    int res = service.deleteById(id);
                    if (res > 0) return Ok();
                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }


        private string generateJwtToken(Customer customer)
        {
            var claims = new Claim[]
            {
                new Claim("Id", customer.Id.ToString()),
                new Claim("Username", customer.Username),
                new Claim(ClaimTypes.Email, customer.Email),
                new Claim("Role", customer.Role)
            };

            // generate token that is valid for 7 days
            var tokenHandler = new JwtSecurityTokenHandler();
            var key = Encoding.ASCII.GetBytes(_appSettings.Secret);
            var tokenDescriptor = new SecurityTokenDescriptor
            {
                Subject = new ClaimsIdentity(claims),
                Expires = DateTime.UtcNow.AddMinutes(30),
                SigningCredentials = new SigningCredentials(new SymmetricSecurityKey(key), SecurityAlgorithms.HmacSha256Signature)
            };
            var token = tokenHandler.CreateToken(tokenDescriptor);
            return tokenHandler.WriteToken(token);
        }

        private string getCustomerRole(string jwtCustomer)
        {
            var handler = new JwtSecurityTokenHandler();
            var jsonToken = handler.ReadToken(jwtCustomer);
            var tokenS = jsonToken as JwtSecurityToken;
            var role = tokenS.Claims.First(claim => claim.Type == "Role").Value;
            return role;
        }
        private string getCustomerId(string jwt)
        {
            var handler = new JwtSecurityTokenHandler();
            var jsonToken = handler.ReadToken(jwt);
            var tokenS = jsonToken as JwtSecurityToken;
            var id = tokenS.Claims.First(claim => claim.Type == "Id").Value;
            return id;
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
