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
    public class ManagerController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly ManagerService service;
        private readonly Context context;
        public ManagerController(IConfiguration config, Context ctx)
        {
            _config = config;
            context = ctx;
            service = new ManagerService(config, context);
        }

        [Route("managers")]
        [HttpGet]
        public JsonResult Get()
        {
            var managers = service.GetAllManagers();
            return new JsonResult(managers);
        }

        [Route("manager/{id}")]
        [HttpGet]
        public ActionResult GetById(int id)
        {
            var manager = service.GetManagerById(id);
            if (manager == null) return BadRequest();
            return new JsonResult(manager);
        }

        [Route("manager/create")]
        [HttpPost]
        public ActionResult Post(Manager manager)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1")
                {
                    if (ModelState.IsValid)
                    {
                        if (service.Post(manager) > 0)
                        {
                            return Ok();
                        }
                    }
                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("managers/export")]
        [HttpGet]
        public ActionResult ExportExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var data = new DataTable();
            var stream = new MemoryStream();
            int count;
            var currentDate = DateTime.Now;
            using (var package = new ExcelPackage(stream))
            {
                var sheet = package.Workbook.Worksheets.Add("Danh sach quan ly vien");
                sheet.Cells["A1:W99"].Style.Font.Name = "Times New Roman";
                sheet.Cells["A1:W99"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

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

                sheet.Cells["K2:N2"].Merge = true;
                sheet.Cells["K2:N2"].Value = "BỘ PHẬN NHÂN SỰ";
                sheet.Cells["K2:N2"].Style.Font.UnderLine = true;
                sheet.Cells["K2:N2"].Style.Font.Size = 13;
                sheet.Cells["K2:N2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["K2:N2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells["K3:N3"].Merge = true;
                sheet.Cells["K3:N3"].Value = $"Ngày xuất: {currentDate.Day}/{currentDate.Month}/{currentDate.Year} " +
                                                $"{currentDate.Hour}:{currentDate.Minute}:{currentDate.Second}";
                sheet.Cells["K3:N3"].Style.Font.Italic = true;
                sheet.Cells["K3:N3"].Style.Font.Size = 13;
                sheet.Cells["K3:N3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["K3:N3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells["A5:N5"].Merge = true;
                sheet.Cells["A5:N5"].Value = "DANH SÁCH QUẢN LÝ VIÊN";
                sheet.Cells["A5:N5"].Style.Font.Bold = true;
                sheet.Cells["A5:N5"].Style.Font.Size = 20;
                sheet.Cells["A5:N5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A5:N5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Row(5).Height = 40;

                sheet.Row(6).Style.WrapText = true;
                sheet.Row(6).Height = 30;

                sheet.Cells["A6:N6"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:N6"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:N6"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:N6"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:N6"].Style.Font.Size = 13;
                sheet.Cells["A6:N6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A6:N6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

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
                sheet.Column(9).Width = 20;
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
                sheet.Cells["L6:L6"].Value = "Quản lý cửa hàng";
                sheet.Cells["L6:L6"].Style.Font.Bold = true;
                sheet.Cells["L6:L6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["L6:L6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);
                sheet.Column(13).Width = 20;
                sheet.Cells["M6:M6"].Value = "Bonus";
                sheet.Cells["M6:M6"].Style.Font.Bold = true;
                sheet.Cells["M6:M6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["M6:M6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);
                sheet.Column(14).Width = 10;
                sheet.Cells["N6:N6"].Value = "Trạng thái";
                sheet.Cells["N6:N6"].Style.Font.Bold = true;
                sheet.Cells["N6:N6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["N6:N6"].Style.Fill.BackgroundColor.SetColor(0, 91, 155, 213);

                var managers = service.GetAllManagers();
                count = managers.Count;
                if (count < 1)
                    return Content("Không có nhân viên nào");

                var stores = context.Stores.ToList<Store>();

                int dem = 1;
                int dong = 7;
                count = count + dong - 1;
                string table = $"A7:N{count}";
                if (count >= 1)
                {
                    sheet.Cells[table].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.WrapText = true;

                    foreach (var manager in managers)
                    {
                        string stt = $"A{dong}:A{dong}";
                        string maqlv = $"B{dong}:B{dong}";
                        string tenqlv = $"C{dong}:C{dong}";
                        string tuoi = $"D{dong}:D{dong}";
                        string gioitinh = $"E{dong}:E{dong}";
                        string email = $"F{dong}:F{dong}";
                        string sdt = $"G{dong}:G{dong}";
                        string diachi = $"H{dong}:H{dong}";
                        string thanhpho = $"I{dong}:I{dong}";
                        string quoctich = $"J{dong}:J{dong}";
                        string luong = $"K{dong}:K{dong}";
                        string tench = $"L{dong}:L{dong}";
                        string bonus = $"M{dong}:M{dong}";
                        string trangthai = $"N{dong}:N{dong}";

                        sheet.Cells[stt].Value = dem;
                        sheet.Cells[maqlv].Value = manager.Id.ToString();
                        sheet.Cells[tenqlv].Value = manager.Name;
                        sheet.Cells[tuoi].Value = manager.Age;
                        sheet.Cells[tuoi].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[gioitinh].Value = manager.Gender == 1 ? "Nam" : "Nữ";
                        sheet.Cells[gioitinh].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[email].Value = manager.Email;
                        sheet.Cells[sdt].Value = manager.Phone.ToString();
                        sheet.Cells[diachi].Value = manager.Address;
                        sheet.Cells[thanhpho].Value = manager.City;
                        sheet.Cells[quoctich].Value = manager.Country;
                        sheet.Cells[quoctich].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[luong].Value = manager.Salary.ToString("C2", CultureInfo.CreateSpecificCulture("vi-VN"));
                        sheet.Cells[luong].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                        foreach (var store in stores)
                        {
                            if (store.ManagerId == manager.Id)
                                sheet.Cells[tench].Value = store.StoreName;
                            else
                                sheet.Cells[tench].Value = "";
                        }
                        sheet.Cells[bonus].Value = manager.Bonus.ToString("C2", CultureInfo.CreateSpecificCulture("vi-VN"));
                        sheet.Cells[bonus].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[trangthai].Value = manager.Status;
                        sheet.Cells[trangthai].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        dong++;
                        dem++;
                    }

                    sheet.Cells[$"K{dong + 1}:N{dong + 1}"].Merge = true;
                    sheet.Cells[$"K{dong + 1}:N{dong + 1}"].Value = $"Hồ Chí Minh, " +
                                                                    $"ngày {currentDate.Day} tháng {currentDate.Month} năm {currentDate.Year}";
                    sheet.Cells[$"K{dong + 1}:N{dong + 1}"].Style.Font.Italic = true;
                    sheet.Cells[$"K{dong + 1}:N{dong + 1}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    sheet.Cells[$"K{dong + 2}:N{dong + 2}"].Merge = true;
                    sheet.Cells[$"K{dong + 2}:N{dong + 2}"].Value = "TRƯỞNG QUẢN LÝ";
                    sheet.Cells[$"K{dong + 2}:N{dong + 2}"].Style.Font.Bold = true;
                    sheet.Cells[$"K{dong + 2}:N{dong + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    sheet.Cells[$"K{dong + 9}:N{dong + 9}"].Merge = true;
                    sheet.Cells[$"K{dong + 9}:N{dong + 9}"].Value = "Võ Hoàng Nhật";
                    sheet.Cells[$"K{dong + 9}:N{dong + 9}"].Style.Font.Bold = true;
                    sheet.Cells[$"K{dong + 9}:N{dong + 9}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }
                package.Save();
            }
            stream.Position = 0;

            var tenfile = $"Danh-sach-quan-ly-vien_{DateTime.Now}.xlsx";

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", tenfile);

        }

        [Route("manager/update/{id}")]
        [HttpPut]
        public ActionResult Put(int id, Manager manager)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1")
                {
                    if (ModelState.IsValid)
                    {
                        if (service.Put(id, manager) > 0)
                            return Ok();
                    }
                    return BadRequest();
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("manager/delete/{id}")]
        [HttpDelete]
        public ActionResult Delete(int id)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var Role = getCurrentRole(jwt);
                if (Role == "1")
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
