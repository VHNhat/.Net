/*using CoffeeBook.Controllers;
using CoffeeBook.JWT;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Net;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Http;

namespace CoffeeBook.Middleware
{
    public class AuthenticationMiddleware
    {
        private readonly RequestDelegate _next;

        public AuthenticationMiddleware(RequestDelegate next)
        {
            _next = next;
        }

        public Task Invoke(HttpContext httpContext, JwtService _jwtService)
        {
            var path = httpContext.Request.Path;
            Console.WriteLine(path);
            string[] adminPath = { "/account" };
            string[] managerPath = { "/account" };

            httpContext.Request.Headers.TryGetValue("Authorization", out var headerValue);
            //Console.WriteLine(headerValue);

            if (headerValue == "Not Author")
            {
                httpContext.Response.StatusCode=401;
            }
            else if(headerValue.ToString() != "" && headerValue.ToString().StartsWith("Bearer"))
            {
                *//*var jwt = headerValue.ToString().Split(" ");
                var token = _jwtService.validate(jwt[1]);
                var handler = new JwtSecurityTokenHandler();
                var jsonToken = handler.ReadToken(jwt[1]);
                var tokenS = jsonToken as JwtSecurityToken;
                var role = tokenS.Claims.First(claim => claim.Type == "RoleId").Value;
                Console.WriteLine(role);
                switch (role)
                {
                    case "1":
                        if (adminPath.Any(a => a.StartsWith(path)))
                        {
                            return _next(httpContext);
                        }
                        else
                        {
                            httpContext.Response.StatusCode = 401;
                            throw new HttpResponseException(HttpStatusCode.Unauthorized);
                        }
                        break;
                    case "2":
                        if (adminPath.Any(a => a.StartsWith(path)))
                        {
                            return _next(httpContext);
                        }
                        else
                        {
                            httpContext.Response.StatusCode = 401;
                            throw new HttpResponseException(HttpStatusCode.Unauthorized);
                        }
                        break;
                    default:
                        httpContext.Response.StatusCode = 401;
                        throw new HttpResponseException(HttpStatusCode.Unauthorized);
                        break;
                }*//*
                //var identity = httpContext.User.Identity as ClaimsIdentity;
                //var claims = token.Claims;
                //var issue = token.Issuer;
                //Console.WriteLine(claims);


            }
            
            return _next(httpContext);
        }
    }

    // Extension method used to add the middleware to the HTTP request pipeline.
    public static class AuthenticationMiddlewareExtensions
    {
        public static IApplicationBuilder UseAuthenticationMiddleware(this IApplicationBuilder builder)
        {
            return builder.UseMiddleware<AuthenticationMiddleware>();
        }
    }

    // Extension method used to add the middleware to the HTTP request pipeline.


}
*/