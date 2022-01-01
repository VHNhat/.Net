using CoffeeBook.DataAccess;
using CoffeeBook.Models;
using CoffeeBook.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IdentityModel.Tokens.Jwt;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Controllers
{
    //[Route("api/[controller]")]
    [ApiController]
    public class EmployeeController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly EmployeeService service;
        private readonly Context context;

        public EmployeeController(IConfiguration config, Context ctx)
        {
            _config = config;
            context = ctx;
            service = new EmployeeService(_config, context);
        }

        [Route("employees")]
        [HttpGet]
        public JsonResult Get()
        {
            var employees = service.GetAllEmployees();
            return new JsonResult(employees);
        }

        [Route("employee/{id}")]
        [HttpGet]
        public ActionResult GetById(int id)
        {
            var employee = service.GetEmployeeById(id);
            if (employee == null) return BadRequest();
            return new JsonResult(employee);
        }

        [Route("employee/create")]
        [HttpPost]
        public ActionResult Post(Employee employee)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1" || Role == "2")
                {
                    if (ModelState.IsValid)
                    {
                        if (service.Post(employee) > 0)
                        {
                            return Ok();
                        }
                    }
                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("employee/export")]
        public ActionResult ExportExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var data = new DataTable();
            var stream = new MemoryStream();
            int count;
            var currentDate = DateTime.Now;
            using (var package = new ExcelPackage(stream))
            {
                var sheet = package.Workbook.Worksheets.Add("Danh sach nhan vien");
                sheet.Cells["A1:W99"].Style.Font.Name = "Times New Roman";

                sheet.Cells["A1:E1"].Merge = true;
                sheet.Cells["A1:E1"].Value = "COFFEE & BOOK";
                sheet.Cells["A1:E1"].Style.Font.Bold = true;
                sheet.Cells["A1:E1"].Style.Font.Size = 14;
                sheet.Cells["A1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                sheet.Cells["A2:C2"].Merge = true;
                sheet.Cells["A2:C2"].Value = "Trụ sở chính: Thành phố HCM";
                sheet.Cells["A2:C2"].Style.Font.Size = 13;
                sheet.Cells["A2:C2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                sheet.Cells["A3:C3"].Merge = true;
                sheet.Cells["A3:C3"].Value = "SĐT: 0901234567";
                sheet.Cells["A3:C3"].Style.Font.Size = 13;
                sheet.Cells["A3:C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                sheet.Cells["I2:L2"].Merge = true;
                sheet.Cells["I2:L2"].Value = "BỘ PHẬN NHÂN SỰ";
                sheet.Cells["I2:L2"].Style.Font.UnderLine = true;
                sheet.Cells["I2:L2"].Style.Font.Size = 13;
                sheet.Cells["I2:L2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["I2:L2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells["I3:L3"].Merge = true;
                sheet.Cells["I3:L3"].Value = $"Ngày xuất: {currentDate.Day}/{currentDate.Month}/{currentDate.Year} " +
                                                $"{currentDate.Hour}:{currentDate.Minute}:{currentDate.Second}";
                sheet.Cells["I3:L3"].Style.Font.Italic = true;
                sheet.Cells["I3:L3"].Style.Font.Size = 13;
                sheet.Cells["I3:L3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["I3:L3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells["A5:M5"].Merge = true;
                sheet.Cells["A5:M5"].Value = "DANH SÁCH NHÂN VIÊN";
                sheet.Cells["A5:M5"].Style.Font.Bold = true;
                sheet.Cells["A5:M5"].Style.Font.Size = 20;
                sheet.Cells["A5:M5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A5:M5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Row(5).Height = 40;

                sheet.Row(6).Style.WrapText = true;
                sheet.Row(6).Height = 30;

                sheet.Cells["A6:M6"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:M6"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:M6"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:M6"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:M6"].Style.Font.Size = 13;
                sheet.Cells["A6:M6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A6:M6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Column(1).Width = 7;
                sheet.Cells["A6:A6"].Value = "STT";
                sheet.Cells["A6:A6"].Style.Font.Bold = true;
                sheet.Cells["A6:A6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["A6:A6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);
                sheet.Column(2).Width = 12;
                sheet.Cells["B6:B6"].Value = "Mã nhân viên";
                sheet.Cells["B6:B6"].Style.Font.Bold = true;
                sheet.Cells["B6:B6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["B6:B6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);
                sheet.Column(3).Width = 20;
                sheet.Cells["C6:C6"].Value = "Tên nhân viên";
                sheet.Cells["C6:C6"].Style.Font.Bold = true;
                sheet.Cells["C6:C6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["C6:C6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);
                sheet.Column(4).Width = 8;
                sheet.Cells["D6:D6"].Value = "Tuổi";
                sheet.Cells["D6:D6"].Style.Font.Bold = true;
                sheet.Cells["D6:D6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["D6:D6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);
                sheet.Column(5).Width = 10;
                sheet.Cells["E6:E6"].Value = "Giới tính";
                sheet.Cells["E6:E6"].Style.Font.Bold = true;
                sheet.Cells["E6:E6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["E6:E6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);
                sheet.Column(6).Width = 25;
                sheet.Cells["F6:F6"].Value = "Email";
                sheet.Cells["F6:F6"].Style.Font.Bold = true;
                sheet.Cells["F6:F6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["F6:F6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);
                sheet.Column(7).Width = 20;
                sheet.Cells["G6:G6"].Value = "Số điện thoại";
                sheet.Cells["G6:G6"].Style.Font.Bold = true;
                sheet.Cells["G6:G6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["G6:G6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);
                sheet.Column(8).Width = 30;
                sheet.Cells["H6:H6"].Value = "Địa chỉ";
                sheet.Cells["H6:H6"].Style.Font.Bold = true;
                sheet.Cells["H6:H6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["H6:H6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);
                sheet.Column(9).Width = 15;
                sheet.Cells["I6:I6"].Value = "Thành phố";
                sheet.Cells["I6:I6"].Style.Font.Bold = true;
                sheet.Cells["I6:I6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["I6:I6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);
                sheet.Column(10).Width = 15;
                sheet.Cells["J6:J6"].Value = "Quốc tịch";
                sheet.Cells["J6:J6"].Style.Font.Bold = true;
                sheet.Cells["J6:J6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["J6:J6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);
                sheet.Column(11).Width = 20;
                sheet.Cells["K6:K6"].Value = "Lương/tháng";
                sheet.Cells["K6:K6"].Style.Font.Bold = true;
                sheet.Cells["K6:K6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["K6:K6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);
                sheet.Column(12).Width = 20;
                sheet.Cells["L6:L6"].Value = "Cửa hàng làm việc";
                sheet.Cells["L6:L6"].Style.Font.Bold = true;
                sheet.Cells["L6:L6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["L6:L6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);
                sheet.Column(13).Width = 10;
                sheet.Cells["M6:M6"].Value = "Trạng thái";
                sheet.Cells["M6:M6"].Style.Font.Bold = true;
                sheet.Cells["M6:M6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["M6:M6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);

                var emps = service.GetAllEmployees();
                count = emps.Count;
                if (count < 1)
                    return Content("Không có nhân viên nào");

                int dem = 1;
                int dong = 7;
                count = count + dong - 1;
                string table = $"A7:M{count}";
                if (count >= 1)
                {
                    sheet.Cells[table].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[table].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    sheet.Cells[table].Style.WrapText = true;

                    foreach (var emp in emps)
                    {
                        string stt = $"A{dong}:A{dong}";
                        string manv = $"B{dong}:B{dong}";
                        string tennv = $"C{dong}:C{dong}";
                        string tuoi = $"D{dong}:D{dong}";
                        string gioitinh = $"E{dong}:E{dong}";
                        string email = $"F{dong}:F{dong}";
                        string sdt = $"G{dong}:G{dong}";
                        string diachi = $"H{dong}:H{dong}";
                        string thanhpho = $"I{dong}:I{dong}";
                        string quoctich = $"J{dong}:J{dong}";
                        string luong = $"K{dong}:K{dong}";
                        string tench = $"L{dong}:L{dong}";
                        string trangthai = $"M{dong}:M{dong}";

                        sheet.Cells[stt].Value = dem;
                        sheet.Cells[manv].Value = emp.Id;
                        sheet.Cells[tennv].Value = emp.Name;
                        sheet.Cells[tuoi].Value = emp.Age;
                        sheet.Cells[gioitinh].Value = emp.Gender == 1 ? "Nam" : "Nữ";
                        sheet.Cells[email].Value = emp.Email;
                        sheet.Cells[sdt].Value = emp.Phone;
                        sheet.Cells[diachi].Value = emp.Address;
                        sheet.Cells[thanhpho].Value = emp.City;
                        sheet.Cells[quoctich].Value = emp.Country;
                        sheet.Cells[luong].Value = emp.Salary.ToString("C2", CultureInfo.CreateSpecificCulture("vi-VN"));
                        var store = context.Stores.Find(emp.StoreId);
                        if (store != null)
                            sheet.Cells[tench].Value = store.StoreName;
                        else
                            sheet.Cells[tench].Value = "";
                        sheet.Cells[trangthai].Value = emp.Status;

                        dong++;
                        dem++;
                    }

                    sheet.Cells[$"I{dong + 1}:M{dong + 1}"].Merge = true;
                    sheet.Cells[$"I{dong + 1}:M{dong + 1}"].Value = $"Hồ Chí Minh, " +
                                                                    $"ngày {currentDate.Day} tháng {currentDate.Month} năm {currentDate.Year}";
                    sheet.Cells[$"I{dong + 1}:M{dong + 1}"].Style.Font.Italic = true;
                    sheet.Cells[$"I{dong + 1}:M{dong + 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[$"I{dong + 1}:M{dong + 1}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    sheet.Cells[$"J{dong + 2}:L{dong + 2}"].Merge = true;
                    sheet.Cells[$"J{dong + 2}:L{dong + 2}"].Value = "QUẢN LÝ VIÊN";
                    sheet.Cells[$"J{dong + 2}:L{dong + 2}"].Style.Font.Bold = true;
                    sheet.Cells[$"J{dong + 2}:L{dong + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[$"J{dong + 2}:L{dong + 2}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    sheet.Cells[$"J{dong + 9}:L{dong + 9}"].Merge = true;
                    sheet.Cells[$"J{dong + 9}:L{dong + 9}"].Value = "Võ Hoàng Nhật";
                    sheet.Cells[$"J{dong + 9}:L{dong + 9}"].Style.Font.Bold = true;
                    sheet.Cells[$"J{dong + 9}:L{dong + 9}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    sheet.Cells[$"J{dong + 9}:L{dong + 9}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }
                package.Save();
            }
            stream.Position = 0;

            var tenfile = $"Danh-sach-nhan-vien_{DateTime.Now}.xlsx";

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", tenfile);
        }

        [Route("employee/update/{id}")]
        [HttpPut]
        public ActionResult Put(int id, Employee employee)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1" || Role == "2")
                {
                    if (ModelState.IsValid)
                    {
                        if (service.Put(id, employee) > 0)
                            return Ok();
                    }
                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("employee/delete/{id}")]
        [HttpDelete]
        public ActionResult Delete(int id)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1" || Role == "2")
                {
                    if (service.Delete(id) > 0)
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
