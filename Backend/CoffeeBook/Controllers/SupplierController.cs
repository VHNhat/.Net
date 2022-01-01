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
    public class SupplierController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly SupplierService service;
        private readonly Context context;

        public SupplierController(IConfiguration config, Context ctx)
        {
            _config = config;
            context = ctx;
            service = new SupplierService(_config, ctx);
        }


        [Route("supplier")]
        [HttpGet]
        public JsonResult Get()
        {
            List<Supplier> suppliers = service.findAll();
            return new JsonResult(suppliers);
        }

        [Route("supplier/{id}")]
        [HttpGet]
        public ActionResult GetById(int id)
        {
            Supplier sup = service.findById(id);
            if (sup == null) return BadRequest();
            else return new JsonResult(sup);
        }

        [Route("supplier/add")]
        [HttpPost]
        public ActionResult Post(Supplier supplier)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1")
                {
                    int res = service.save(supplier);
                    if (res > 0) return Ok();
                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("supplier/edit/{id}")]
        [HttpPut]
        public ActionResult Put(int id,Supplier supplier)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1")
                {
                    int res = service.Update(id, supplier);
                    if (res > 0) return Ok();
                    else return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("supplier/delete/{id}")]
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
