using CoffeeBook.Authen;
using CoffeeBook.DataAccess;
using CoffeeBook.Dto;
using CoffeeBook.Models;
using CoffeeBook.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Tokens;
using System;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;

namespace CoffeeBook.Controllers
{
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
            var customers = service.FindAll();
            if (customers == null || customers.Count == 0)
                return BadRequest();
            
            return new JsonResult(customers);
        }

        [Route("customer/{id}")]
        [HttpGet]
        public ActionResult Get(int id)
        {
            var customer = service.FindById(id);
            if (customer == null) 
                return BadRequest();

            return new JsonResult(customer);
        }

        [Route("customer/add")]
        [HttpPost]
        public ActionResult Post(Customer customer)
        {
            var result = service.Add(customer);
            if (result > 0)
                return Ok();

            return BadRequest();
        }

        [Route("customer/login")]
        [HttpPost]
        public ActionResult Login(SigninDto dto)
        {
            Customer user = service.Login(dto);

            if(user == null)
                return BadRequest();

            var token = GenerateJwtToken(user);

            return new JsonResult(new { Token = token });
        }

        [Route("customer/signup")]
        [HttpPost]
        public ActionResult Register(SignupDto dto)
        {
            var result = service.Register(dto);
            if (result == "1")
                return Ok();

            return new JsonResult(result);
        }

        [Route("customer/edit/{id}")]
        [HttpPut]
        public ActionResult Put(int id,Customer customer)
        {
            int res = service.Update(id,customer);
            if (res > 0) 
                return Ok();
            
            return BadRequest();
        }

        [Route("customer/delete/{id}")]
        [HttpDelete]
        public ActionResult Delete(int id)
        {
            var result = service.Delete(id);
            if (result > 0)
                return Ok();

            return BadRequest();
        }

        private string GenerateJwtToken(Customer customer)
        {
            var claims = new Claim[]
            {
                new Claim("Id", customer.Id.ToString()),
                new Claim("Username", customer.Username),
                new Claim(ClaimTypes.Email, customer.Email)
            };

            // generate token that is valid for 7 days
            var tokenHandler = new JwtSecurityTokenHandler();
            var key = Encoding.ASCII.GetBytes(_appSettings.Secret);
            var tokenDescriptor = new SecurityTokenDescriptor
            {
                Subject = new ClaimsIdentity(claims),
                Expires = DateTime.UtcNow.AddSeconds(20),
                SigningCredentials = new SigningCredentials(new SymmetricSecurityKey(key), SecurityAlgorithms.HmacSha256Signature)
            };
            var token = tokenHandler.CreateToken(tokenDescriptor);
            return tokenHandler.WriteToken(token);
        }
    }
}
