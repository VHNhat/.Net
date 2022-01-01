using CoffeeBook.DataAccess;
using CoffeeBook.Models;
using CoffeeBook.Services;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Controllers
{
    //[Route("api/[controller]")]
    [ApiController]
    public class RoleController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly RoleService _service;
        private readonly Context _context;
        public RoleController(IConfiguration config, Context context)
        {
            _config = config;
            _context = context;
            _service = new RoleService(_config, _context);
        }
        [Route("role")]
        [HttpGet]
        public ActionResult Get()
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1")
                {
                    var roles = _service.GetAllRoles();
                    return new JsonResult(roles);
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("role/{id}")]
        [HttpGet]
        public ActionResult GetById(int id)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1")
                {
                    Role role = _service.GetRoleById(id);
                    if (role == null) return BadRequest();
                    else return new JsonResult(role);
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("role/add")]
        [HttpPost]
        public ActionResult AddRole(Role role)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1")
                {
                    int result = _service.Post(role);
                    if (result > 0)
                        return Ok();
                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("role/edit/{id}")]
        [HttpPut]
        public IActionResult UpdateRole(int id, Role role)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1")
                {
                    if (ModelState.IsValid)
                    {
                        int res = _service.Put(id, role);
                        if (res > 0) return Ok();
                        else return BadRequest();
                    }
                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("role/delete/{id}")]
        [HttpDelete]
        public IActionResult DeleteRole(int id)
        {
            
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1")
                {
                    int res = _service.Delete(id);
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
