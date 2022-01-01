using CoffeeBook.DataAccess;
using CoffeeBook.Models;
using CoffeeBook.Services;
using Microsoft.AspNetCore.Authorization;
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
    public class NewsController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly NewsService service;
        private readonly Context context;

        public NewsController(IConfiguration config, Context ctx)
        {
            _config = config;
            context = ctx;
            service = new NewsService(_config, ctx);
        }


        [Route("news")]
        [HttpGet]
        public ActionResult Get()
        {
            List<News> news = service.findAll();
            return new JsonResult(news);
        }

        [Route("news/{id}")]
        [HttpGet]
        public ActionResult Get(int id)
        {
            News news = service.FindById(id);
            if (news == null)
                return BadRequest();

            return new JsonResult(news);
        }

        [Authorize(Roles = "1")]
        [Route("news/add")]
        [HttpPost]
        public ActionResult Post(News news)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1" || Role == "2")
                {
                    int res = service.save(news);
                    if (res > 0) return Ok();

                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("news/edit/{id}")]
        [HttpPut]
        public ActionResult Put(int id, News news)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1" || Role == "2")
                {
                    int res = service.update(id, news);
                    if (res > 0) return Ok();
                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("news/delete/{id}")]
        [HttpDelete]
        public ActionResult Delete(int id)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1" || Role == "2")
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
