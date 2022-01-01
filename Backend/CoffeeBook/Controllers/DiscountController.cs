using CoffeeBook.DataAccess;
using CoffeeBook.Models;
using CoffeeBook.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Controllers
{
    /*[Route("api/[controller]")]*/
    [ApiController]
    public class DiscountController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly DiscountService service;
        private readonly Context context;

        public DiscountController(IConfiguration config, Context ctx)
        {
            _config = config;
            context = ctx;
            service = new DiscountService(_config, ctx);
        }


        [Route("discount")]
        [HttpGet]
        public ActionResult Get()
        {
            var discounts = service.FindAll();
            return new JsonResult(discounts);
        }

        [Route("discount/{id}")]
        [HttpGet]
        public ActionResult Get(int id)
        {
            var discount = service.FindById(id);
            if (discount == null)
                return BadRequest();

            return new JsonResult(discount);
        }

        [Route("discount/add")]
        [HttpPost]
        public ActionResult Post(Discount discount)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1" || Role == "2")
                {
                    var result = service.save(discount);
                    if (result > 0)
                        return Ok();

                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("discount/edit/{id}")]
        [HttpPut]
        public ActionResult Put(int id, Discount discount)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1" || Role == "2")
                {
                    var result = service.Update(id, discount);

                    if (result > 0)
                        return Ok();

                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("discount/delete/{id}")]
        [HttpDelete]
        public ActionResult Delete(int id)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1" || Role == "2")
                {
                    var result = service.DeleteById(id);

                    if (result > 0)
                        return Ok();

                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
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
