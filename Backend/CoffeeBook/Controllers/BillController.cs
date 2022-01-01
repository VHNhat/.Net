using CoffeeBook.DataAccess;
using CoffeeBook.Dto;
using CoffeeBook.Models;
using CoffeeBook.Services;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using MySql.Data.MySqlClient;
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
    /*[Route("api/[controller]")]*/
    [ApiController]
    public class BillController : ControllerBase
    {
        private readonly IConfiguration _config;
        private readonly Context context;
        private readonly BillService _service;
        public BillController(IConfiguration config, Context ctx)
        {
            _config = config;
            context = ctx;
            _service = new BillService(_config, ctx);

        }

        [Route("bill")]
        [HttpGet]
        public ActionResult Get()
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var role = getCurrentRole(jwt);
                if (role == "1" || role == "2" || role == "3" || role == "4")
                {
                    List<Bill> bills = _service.findAll();
                    return new JsonResult(bills);
                }
            }
            return Unauthorized(new { message = "Bạn không có quyền truy cập" });
        }

        [Route("bill/{id}")]
        [HttpGet]
        public JsonResult GetById(int id)
        {
            List<Bill> bills = _service.GetBillId(id);
            return new JsonResult(bills);
        }
        [Route("bill/sales/year/{year}")]
        [HttpGet]
        public JsonResult GetSaleByYearMonth(int year)
        {
            var sales = _service.GetSaleByYear(year);
            return new JsonResult(sales);
        }

        [Route("bill/totalsale/{year}")]
        [HttpGet]
        public ActionResult GetTotalSaleByYear(int year)
        {
            var sale = _service.GetTotalSaleByYear(year);
            return new JsonResult(sale);
        }
        [Route("bill/count/{year}")]
        public ActionResult GetCountBill(int year)
        {
            var count = _service.GetCountBill(year);
            return new JsonResult(count );
        }
        // url: .../bill/sales/date/month-day-year
        // example: .../bill/sales/date/12-24-2021
        [Route("bill/sales/date/{date}")]
        [HttpGet]
        public ActionResult GetSaleByDay(DateTime date)
        {

            var sale = _service.GetRevenueByDay(date);
            return new JsonResult(sale);
        }
      ////
        [Route("bill/sales/months/{param}")]
        [HttpGet]
        public ActionResult GetSaleByMonth(string param)
        {
            string[] parameters = param.Split("-");
            int month = int.Parse(parameters[0]);
            int year = int.Parse(parameters[1]);
            var sale = _service.GetRevenueByMonth(month, year);
            return new JsonResult(sale);
        }
        /// <summary>
        ///
        /// <returns></returns>
        [Route("bill/sales/years/{year}")]
        [HttpGet]
        public ActionResult GetSaleByYear(int year)
        {
            var sale = _service.GetRevenueByYear(year);
            return new JsonResult(sale);
        }

        [Route("bill/total/{year}")]
        [HttpGet]
        public ActionResult GetTotalListBill(int year)
        {
            var count = _service.GetTotalListBill(year);
            return new JsonResult( count);
        }

        // url: .../bill/revenue/date/month-day-year
        // example: .../bill/revenue/date/12-24-2021
        [Route("bill/revenue/date/{date}")]
        [HttpGet]
        public ActionResult GetRevenueByDay(DateTime date)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var data = new DataTable();
            var stream = new MemoryStream();
            int count;
            var currentDate = DateTime.Now;

            using (var package = new ExcelPackage(stream))
            {
                string sheetName = $"Doanh thu ngay {date.Day}/{date.Month}/{date.Year}";
                var sheet = package.Workbook.Worksheets.Add(sheetName);
                sheet.Cells["A1:W99"].Style.Font.Name = "Times New Roman";

                sheet.Cells["A1:F1"].Merge = true;
                sheet.Cells["A1:F1"].Value = "COFFEE & BOOK";
                sheet.Cells["A1:F1"].Style.Font.Bold = true;
                sheet.Cells["A1:F1"].Style.Font.Size = 14;
                sheet.Cells["A1:F1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells["A1:F1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells["A2:D2"].Merge = true;
                sheet.Cells["A2:D2"].Value = "Trụ sở chính: Thành phố HCM";
                sheet.Cells["A2:D2"].Style.Font.Size = 13
                    ;
                sheet.Cells["A2:D2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells["A2:D2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["A3:C3"].Merge = true;
                sheet.Cells["A3:C3"].Value = "SĐT: 0901234567";
                sheet.Cells["A3:C3"].Style.Font.Size = 13;
                sheet.Cells["A3:C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                sheet.Cells["G2:J2"].Merge = true;
                sheet.Cells["G2:J2"].Value = "BỘ PHẬN KẾ TOÁN";
                sheet.Cells["G2:J2"].Style.Font.UnderLine = true;
                sheet.Cells["G2:J2"].Style.Font.Size = 13;
                sheet.Cells["G2:J2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["G2:J2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells["G3:J3"].Merge = true;
                sheet.Cells["G3:J3"].Value = $"Ngày xuất: {currentDate.Day}/{currentDate.Month}/{currentDate.Year} " +
                                                $"{currentDate.Hour}:{currentDate.Minute}:{currentDate.Second}";
                sheet.Cells["G3:J3"].Style.Font.Italic = true;
                sheet.Cells["G3:J3"].Style.Font.Size = 13;
                sheet.Cells["G3:J3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["G3:J3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells["A5:E5"].Merge = true;
                sheet.Cells["A5:E5"].Value = $"THỐNG KÊ DOANH THU NGÀY {date.Day}/{date.Month}/{date.Year}";
                sheet.Cells["A5:E5"].Style.Font.Bold = true;
                sheet.Cells["A5:E5"].Style.Font.Size = 18;
                sheet.Cells["A5:E5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A5:E5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Row(5).Height = 35;

                sheet.Row(6).Style.WrapText = true;
                sheet.Row(6).Height = 30;

                sheet.Cells["A6:E6"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:E6"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:E6"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:E6"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:E6"].Style.Font.Size = 11;
                sheet.Cells["A6:E6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A6:E6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Column(1).Width = 7;
                sheet.Cells["A6:A6"].Value = "STT";
                sheet.Cells["A6:A6"].Style.Font.Bold = true;
                sheet.Cells["A6:A6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["A6:A6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);
                sheet.Column(2).Width = 15;
                sheet.Cells["B6:B6"].Value = "Mã hóa đơn";
                sheet.Cells["B6:B6"].Style.Font.Bold = true;
                sheet.Cells["B6:B6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["B6:B6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);
                sheet.Column(3).Width = 17;
                sheet.Cells["C6:C6"].Value = "Giờ";
                sheet.Cells["C6:C6"].Style.Font.Bold = true;
                sheet.Cells["C6:C6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["C6:C6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);
                sheet.Column(4).Width = 25;
                sheet.Cells["D6:D6"].Value = "Tiền HĐ";
                sheet.Cells["D6:D6"].Style.Font.Bold = true;
                sheet.Cells["D6:D6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["D6:D6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);
                sheet.Column(5).Width = 15;
                sheet.Cells["E6:E6"].Value = "Thanh toán bằng";
                sheet.Cells["E6:E6"].Style.Font.Bold = true;
                sheet.Cells["E6:E6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["E6:E6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);
                sheet.Column(6).Width = 10;
                sheet.Cells["F6:F6"].Value = "Ghi chú";
                sheet.Cells["F6:F6"].Style.Font.Bold = true;
                sheet.Cells["F6:F6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["F6:F6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);

                var revenue = _service.GetRevenueByDay(date);
                int numOfRow = revenue.Rows.Count;
                if (numOfRow < 1)
                    return BadRequest();

                long total = 0;
                int dem = 1;
                int dong = 7;
                count = numOfRow + dong - 1;
                string table = $"A7:F{count}";
                if (count >= 1)
                {
                    sheet.Cells[table].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //revenue = _service.GetRevenueByMonth(month, year);
                    foreach (DataRow row in revenue.Rows)
                    {
                        string stt = $"A{dong}:A{dong}";
                        string maHd = $"B{dong}:B{dong}";
                        string gioHd = $"C{dong}:C{dong}";
                        string tienHd = $"D{dong}:D{dong}";
                        string thanhToan = $"E{dong}:E{dong}";
                        string ghiChu = $"F{dong}:F{dong}";

                        sheet.Cells[stt].Value = dem;
                        sheet.Cells[stt].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[stt].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        sheet.Cells[maHd].Value = row[0].ToString();
                        sheet.Cells[maHd].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[maHd].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        sheet.Cells[gioHd].Value = DateTime.Parse(row[1].ToString()).ToLongTimeString();
                        sheet.Cells[gioHd].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[gioHd].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        total += long.Parse(row[2].ToString());
                        sheet.Cells[tienHd].Value = long.Parse(row[2].ToString()).ToString("C2", CultureInfo.CreateSpecificCulture("vi-VN"));
                        sheet.Cells[tienHd].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[tienHd].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        sheet.Cells[thanhToan].Value = row[3].ToString();
                        sheet.Cells[thanhToan].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[thanhToan].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        sheet.Cells[ghiChu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[ghiChu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        dem++;
                        dong++;
                    }
                }
                sheet.Cells[$"C{dong}:C{dong}"].Value = "Doanh thu:";
                sheet.Cells[$"C{dong}:C{dong}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[$"C{dong}:C{dong}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[$"D{dong}:D{dong}"].Value = total.ToString("C2", CultureInfo.CreateSpecificCulture("vi-VN"));
                sheet.Cells[$"D{dong}:D{dong}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[$"D{dong}:D{dong}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[$"F{dong + 2}:J{dong + 2}"].Merge = true;
                sheet.Cells[$"F{dong + 2}:J{dong + 2}"].Value = $"Hồ Chí Minh, " +
                                                                $"ngày {currentDate.Day} tháng {currentDate.Month} năm {currentDate.Year}";
                sheet.Cells[$"F{dong + 2}:J{dong + 2}"].Style.Font.Italic = true;
                sheet.Cells[$"F{dong + 2}:J{dong + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[$"F{dong + 2}:J{dong + 2}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells[$"G{dong + 3}:I{dong + 3}"].Merge = true;
                sheet.Cells[$"G{dong + 3}:I{dong + 3}"].Value = "KẾ TOÁN TRƯỞNG";
                sheet.Cells[$"G{dong + 3}:I{dong + 3}"].Style.Font.Bold = true;
                sheet.Cells[$"G{dong + 3}:I{dong + 3}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[$"G{dong + 3}:I{dong + 3}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells[$"G{dong + 9}:I{dong + 9}"].Merge = true;
                sheet.Cells[$"G{dong + 9}:I{dong + 9}"].Value = "Võ Hoàng Nhật";
                sheet.Cells[$"G{dong + 9}:I{dong + 9}"].Style.Font.Bold = true;
                sheet.Cells[$"G{dong + 9}:I{dong + 9}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[$"G{dong + 9}:I{dong + 9}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                package.Save();
            }
            stream.Position = 0;

            var tenfile = $"Doanh-thu-ngay-{date.Day}/{date.Month}/{date.Year}_{DateTime.Now.ToString("dd/MM/yyyy")}.xlsx";

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", tenfile);
        }

        // url: ... 
        [Route("bill/revenue/month/{param}")]
        [HttpGet]
        public ActionResult GetRevenueByMonth(string param)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var data = new DataTable();
            var stream = new MemoryStream();
            int count;
            var currentDate = DateTime.Now;

            string[] parameters = param.Split("-");
            int month = int.Parse(parameters[0]);
            int year = int.Parse(parameters[1]);

            using (var package = new ExcelPackage(stream))
            {
                string sheetName = $"Doanh thu thang {month}/{year}";
                var sheet = package.Workbook.Worksheets.Add(sheetName);
                sheet.Cells["A1:W99"].Style.Font.Name = "Times New Roman";

                sheet.Cells["A1:E1"].Merge = true;
                sheet.Cells["A1:E1"].Value = "COFFEE & BOOK";
                sheet.Cells["A1:E1"].Style.Font.Bold = true;
                sheet.Cells["A1:E1"].Style.Font.Size = 14;
                sheet.Cells["A1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells["A1:E1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells["A2:D2"].Merge = true;
                sheet.Cells["A2:D2"].Value = "Trụ sở chính: Thành phố HCM";
                sheet.Cells["A2:D2"].Style.Font.Size = 13;
                sheet.Cells["A2:D2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells["A2:D2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["A3:C3"].Merge = true;
                sheet.Cells["A3:C3"].Value = "SĐT: 0901234567";
                sheet.Cells["A3:C3"].Style.Font.Size = 13;
                sheet.Cells["A3:C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                sheet.Cells["G2:J2"].Merge = true;
                sheet.Cells["G2:J2"].Value = "BỘ PHẬN KẾ TOÁN";
                sheet.Cells["G2:J2"].Style.Font.UnderLine = true;
                sheet.Cells["G2:J2"].Style.Font.Size = 13;
                sheet.Cells["G2:J2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["G2:J2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells["G3:J3"].Merge = true;
                sheet.Cells["G3:J3"].Value = $"Ngày xuất: {currentDate.Day}/{currentDate.Month}/{currentDate.Year} " +
                                                $"{currentDate.Hour}:{currentDate.Minute}:{currentDate.Second}";
                sheet.Cells["G3:J3"].Style.Font.Italic = true;
                sheet.Cells["G3:J3"].Style.Font.Size = 13;
                sheet.Cells["G3:J3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["G3:J3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells["A5:E5"].Merge = true;
                sheet.Cells["A5:E5"].Value = $"THỐNG KÊ DOANH THU THÁNG {month}/{year}";
                sheet.Cells["A5:E5"].Style.Font.Bold = true;
                sheet.Cells["A5:E5"].Style.Font.Size = 18;
                sheet.Cells["A5:E5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A5:E5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Row(5).Height = 35;

                sheet.Row(6).Style.WrapText = true;
                sheet.Row(6).Height = 30;

                sheet.Cells["A6:E6"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:E6"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:E6"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:E6"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:E6"].Style.Font.Size = 11;
                sheet.Cells["A6:E6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A6:E6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Column(1).Width = 7;
                sheet.Cells["A6:A6"].Value = "STT";
                sheet.Cells["A6:A6"].Style.Font.Bold = true;
                sheet.Cells["A6:A6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["A6:A6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);
                sheet.Column(2).Width = 20;
                sheet.Cells["B6:B6"].Value = "Ngày hóa đơn";
                sheet.Cells["B6:B6"].Style.Font.Bold = true;
                sheet.Cells["B6:B6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["B6:B6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);
                sheet.Column(3).Width = 10;
                sheet.Cells["C6:C6"].Value = "Sl hóa đơn";
                sheet.Cells["C6:C6"].Style.Font.Bold = true;
                sheet.Cells["C6:C6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["C6:C6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);
                sheet.Column(4).Width = 25;
                sheet.Cells["D6:D6"].Value = "DThu trong ngày";
                sheet.Cells["D6:D6"].Style.Font.Bold = true;
                sheet.Cells["D6:D6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["D6:D6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);
                sheet.Column(5).Width = 10;
                sheet.Cells["E6:E6"].Value = "Đánh giá";
                sheet.Cells["E6:E6"].Style.Font.Bold = true;
                sheet.Cells["E6:E6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["E6:E6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);

                var revenue = _service.GetRevenueByMonth(month, year);
                int numOfRow = revenue.Rows.Count;
                if (numOfRow < 1)
                    return BadRequest();

                long totalOfMonth = 0;
                int dem = 1;
                int dong = 7;
                count = numOfRow + dong - 1;
                string table = $"A7:E{count}";
                if (count >= 1)
                {
                    sheet.Cells[table].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //revenue = _service.GetRevenueByMonth(month, year);
                    foreach (DataRow row in revenue.Rows)
                    {
                        string stt = $"A{dong}:A{dong}";
                        string ngayHd = $"B{dong}:B{dong}";
                        string sluongHd = $"C{dong}:C{dong}";
                        string doanhThu = $"D{dong}:D{dong}";
                        string ghichu = $"E{dong}:E{dong}";

                        sheet.Cells[stt].Value = dem;
                        sheet.Cells[stt].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[stt].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        sheet.Cells[ngayHd].Value = row[0].ToString() + $"/{month}/{year}";
                        sheet.Cells[ngayHd].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[ngayHd].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        sheet.Cells[sluongHd].Value = int.Parse(row[1].ToString());
                        sheet.Cells[sluongHd].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[sluongHd].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        totalOfMonth += long.Parse(row[2].ToString());
                        sheet.Cells[doanhThu].Value = long.Parse(row[2].ToString()).ToString("C2", CultureInfo.CreateSpecificCulture("vi-VN"));
                        sheet.Cells[doanhThu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[doanhThu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        if (long.Parse(row[2].ToString()) > 2000000)
                        {
                            sheet.Cells[ghichu].Value = "Cao";
                            sheet.Cells[ghichu].Style.Font.Color.SetColor(0, 0, 255, 0);
                        }
                        else if (long.Parse(row[2].ToString()) > 1000000)
                        {
                            sheet.Cells[ghichu].Value = "Trung bình";
                            sheet.Cells[ghichu].Style.Font.Color.SetColor(0, 0, 0, 255);
                        }
                        else
                        {
                            sheet.Cells[ghichu].Value = "Thấp";
                            sheet.Cells[ghichu].Style.Font.Color.SetColor(0, 255, 0, 0);
                        }

                        sheet.Cells[ghichu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[ghichu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        dem++;
                        dong++;
                    }
                }
                sheet.Cells[$"C{dong}:C{dong}"].Value = "Doanh thu:";
                sheet.Cells[$"C{dong}:C{dong}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[$"C{dong}:C{dong}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[$"D{dong}:D{dong}"].Value = totalOfMonth.ToString("C2", CultureInfo.CreateSpecificCulture("vi-VN"));
                sheet.Cells[$"D{dong}:D{dong}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[$"D{dong}:D{dong}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[$"F{dong + 2}:J{dong + 2}"].Merge = true;
                sheet.Cells[$"F{dong + 2}:J{dong + 2}"].Value = $"Hồ Chí Minh, " +
                                                                $"ngày {currentDate.Day} tháng {currentDate.Month} năm {currentDate.Year}";
                sheet.Cells[$"F{dong + 2}:J{dong + 2}"].Style.Font.Italic = true;
                sheet.Cells[$"F{dong + 2}:J{dong + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[$"F{dong + 2}:J{dong + 2}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells[$"G{dong + 3}:I{dong + 3}"].Merge = true;
                sheet.Cells[$"G{dong + 3}:I{dong + 3}"].Value = "KẾ TOÁN TRƯỞNG";
                sheet.Cells[$"G{dong + 3}:I{dong + 3}"].Style.Font.Bold = true;
                sheet.Cells[$"G{dong + 3}:I{dong + 3}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[$"G{dong + 3}:I{dong + 3}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells[$"G{dong + 9}:I{dong + 9}"].Merge = true;
                sheet.Cells[$"G{dong + 9}:I{dong + 9}"].Value = "Võ Hoàng Nhật";
                sheet.Cells[$"G{dong + 9}:I{dong + 9}"].Style.Font.Bold = true;
                sheet.Cells[$"G{dong + 9}:I{dong + 9}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[$"G{dong + 9}:I{dong + 9}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                package.Save();
            }
            stream.Position = 0;

            var tenfile = $"Doanh-thu-thang-{month}-{year}_{DateTime.Now}.xlsx";

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", tenfile);
        }

        // url: .../bill/revenue/year/2021
        [Route("bill/revenue/year/{year}")]
        [HttpGet]
        public ActionResult GetRevenueByYear(int year)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var data = new DataTable();
            var stream = new MemoryStream();
            int count;
            var currentDate = DateTime.Now;

            using (var package = new ExcelPackage(stream))
            {
                string sheetName = $"Doanh thu năm {year}";
                var sheet = package.Workbook.Worksheets.Add(sheetName);
                sheet.Cells["A1:W99"].Style.Font.Name = "Times New Roman";

                sheet.Cells["A1:E1"].Merge = true;
                sheet.Cells["A1:E1"].Value = "COFFEE & BOOK";
                sheet.Cells["A1:E1"].Style.Font.Bold = true;
                sheet.Cells["A1:E1"].Style.Font.Size = 14;
                sheet.Cells["A1:E1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells["A1:E1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells["A2:D2"].Merge = true;
                sheet.Cells["A2:D2"].Value = "Trụ sở chính: Thành phố HCM";
                sheet.Cells["A2:D2"].Style.Font.Size = 13;
                sheet.Cells["A2:D2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                sheet.Cells["A2:D2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells["A3:C3"].Merge = true;
                sheet.Cells["A3:C3"].Value = "SĐT: 0901234567";
                sheet.Cells["A3:C3"].Style.Font.Size = 13;
                sheet.Cells["A3:C3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;

                sheet.Cells["G2:J2"].Merge = true;
                sheet.Cells["G2:J2"].Value = "BỘ PHẬN KẾ TOÁN";
                sheet.Cells["G2:J2"].Style.Font.UnderLine = true;
                sheet.Cells["G2:J2"].Style.Font.Size = 13;
                sheet.Cells["G2:J2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["G2:J2"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells["G3:J3"].Merge = true;
                sheet.Cells["G3:J3"].Value = $"Ngày xuất: {currentDate.Day}/{currentDate.Month}/{currentDate.Year} " +
                                                $"{currentDate.Hour}:{currentDate.Minute}:{currentDate.Second}";
                sheet.Cells["G3:J3"].Style.Font.Italic = true;
                sheet.Cells["G3:J3"].Style.Font.Size = 13;
                sheet.Cells["G3:J3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["G3:J3"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells["A5:E5"].Merge = true;
                sheet.Cells["A5:E5"].Value = $"THỐNG KÊ DOANH THU NĂM {year}";
                sheet.Cells["A5:E5"].Style.Font.Bold = true;
                sheet.Cells["A5:E5"].Style.Font.Size = 18;
                sheet.Cells["A5:E5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A5:E5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Row(5).Height = 35;

                sheet.Row(6).Style.WrapText = true;
                sheet.Row(6).Height = 30;

                sheet.Cells["A6:E6"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:E6"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:E6"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:E6"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells["A6:E6"].Style.Font.Size = 11;
                sheet.Cells["A6:E6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells["A6:E6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Column(1).Width = 7;
                sheet.Cells["A6:A6"].Value = "STT";
                sheet.Cells["A6:A6"].Style.Font.Bold = true;
                sheet.Cells["A6:A6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["A6:A6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);
                sheet.Column(2).Width = 9;
                sheet.Cells["B6:B6"].Value = "Tháng";
                sheet.Cells["B6:B6"].Style.Font.Bold = true;
                sheet.Cells["B6:B6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["B6:B6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);
                sheet.Column(3).Width = 10;
                sheet.Cells["C6:C6"].Value = "SL hóa đơn";
                sheet.Cells["C6:C6"].Style.Font.Bold = true;
                sheet.Cells["C6:C6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["C6:C6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);
                sheet.Column(4).Width = 25;
                sheet.Cells["D6:D6"].Value = "DThu trong tháng";
                sheet.Cells["D6:D6"].Style.Font.Bold = true;
                sheet.Cells["D6:D6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["D6:D6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);
                sheet.Column(5).Width = 10;
                sheet.Cells["E6:E6"].Value = "Đánh giá";
                sheet.Cells["E6:E6"].Style.Font.Bold = true;
                sheet.Cells["E6:E6"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells["E6:E6"].Style.Fill.BackgroundColor.SetColor(0, 169, 208, 142);

                var revenue = _service.GetRevenueByYear(year);
                int numOfRow = revenue.Rows.Count;
                if (numOfRow < 1)
                    return BadRequest();

                long totalOfMonth = 0;
                int dem = 1;
                int dong = 7;
                count = numOfRow + dong - 1;
                string table = $"A7:E{count}";
                if (count >= 1)
                {
                    sheet.Cells[table].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sheet.Cells[table].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    //revenue = _service.GetRevenueByMonth(month, year);
                    foreach (DataRow row in revenue.Rows)
                    {
                        string stt = $"A{dong}:A{dong}";
                        string thangHd = $"B{dong}:B{dong}";
                        string sluongHd = $"C{dong}:C{dong}";
                        string doanhThu = $"D{dong}:D{dong}";
                        string ghichu = $"E{dong}:E{dong}";

                        sheet.Cells[stt].Value = dem;
                        sheet.Cells[stt].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[stt].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        sheet.Cells[thangHd].Value = row[0].ToString();
                        sheet.Cells[thangHd].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[thangHd].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        sheet.Cells[sluongHd].Value = int.Parse(row[1].ToString());
                        sheet.Cells[sluongHd].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[sluongHd].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        totalOfMonth += long.Parse(row[2].ToString());
                        sheet.Cells[doanhThu].Value = long.Parse(row[2].ToString()).ToString("C2", CultureInfo.CreateSpecificCulture("vi-VN"));
                        sheet.Cells[doanhThu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                        sheet.Cells[doanhThu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        if (long.Parse(row[2].ToString()) > 50_000_000)
                        {
                            sheet.Cells[ghichu].Value = "Cao";
                            sheet.Cells[ghichu].Style.Font.Color.SetColor(0, 0, 255, 0);
                        }
                        else if (long.Parse(row[2].ToString()) > 20_000_000)
                        {
                            sheet.Cells[ghichu].Value = "Trung bình";
                            sheet.Cells[ghichu].Style.Font.Color.SetColor(0, 0, 0, 255);
                        }
                        else
                        {
                            sheet.Cells[ghichu].Value = "Thấp";
                            sheet.Cells[ghichu].Style.Font.Color.SetColor(0, 255, 0, 0);
                        }

                        sheet.Cells[ghichu].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        sheet.Cells[ghichu].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        dem++;
                        dong++;
                    }
                }
                sheet.Cells[$"C{dong}:C{dong}"].Value = "Doanh thu:";
                sheet.Cells[$"C{dong}:C{dong}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[$"C{dong}:C{dong}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[$"D{dong}:D{dong}"].Value = totalOfMonth.ToString("C2", CultureInfo.CreateSpecificCulture("vi-VN"));
                sheet.Cells[$"D{dong}:D{dong}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                sheet.Cells[$"D{dong}:D{dong}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                sheet.Cells[$"F{dong + 2}:J{dong + 2}"].Merge = true;
                sheet.Cells[$"F{dong + 2}:J{dong + 2}"].Value = $"Hồ Chí Minh, " +
                                                                $"ngày {currentDate.Day} tháng {currentDate.Month} năm {currentDate.Year}";
                sheet.Cells[$"F{dong + 2}:J{dong + 2}"].Style.Font.Italic = true;
                sheet.Cells[$"F{dong + 2}:J{dong + 2}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[$"F{dong + 2}:J{dong + 2}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells[$"G{dong + 3}:I{dong + 3}"].Merge = true;
                sheet.Cells[$"G{dong + 3}:I{dong + 3}"].Value = "KẾ TOÁN TRƯỞNG";
                sheet.Cells[$"G{dong + 3}:I{dong + 3}"].Style.Font.Bold = true;
                sheet.Cells[$"G{dong + 3}:I{dong + 3}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[$"G{dong + 3}:I{dong + 3}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                sheet.Cells[$"G{dong + 9}:I{dong + 9}"].Merge = true;
                sheet.Cells[$"G{dong + 9}:I{dong + 9}"].Value = "Võ Hoàng Nhật";
                sheet.Cells[$"G{dong + 9}:I{dong + 9}"].Style.Font.Bold = true;
                sheet.Cells[$"G{dong + 9}:I{dong + 9}"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                sheet.Cells[$"G{dong + 9}:I{dong + 9}"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                package.Save();
            }
            stream.Position = 0;

            var tenfile = $"Doanh-thu-nam-{year}_{DateTime.Now}.xlsx";

            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", tenfile);
        }


        [Route("bill/add")]
        [HttpPost]
        public ActionResult Post(Bill bill)
        {
            int table = _service.save(bill);
            if (table > 0) return Ok();
            else return BadRequest();
        }

        [Route("bill/edit/{id}")]
        [HttpPut]
        public ActionResult Put(int id, Bill bill)
        {
            int res = _service.update(id, bill);
            if (res > 0) return Ok();
            else return BadRequest();
        }

        [Route("bill/delete/{id}")]
        [HttpDelete]
        public ActionResult Delete(int id)
        {
            int res = _service.DeleteById(id);
            if (res > 0) return Ok();
            else return BadRequest();
        }

        [Route("bill/purchase")]
        [HttpPost]
        public ActionResult Purchase(BillDto dto)
        {
            int result = _service.Purchase(dto);
            if (result > 0)
                return Ok();
            else return Ok();
        }

        [Route("bill/delivery/{id}")]
        [HttpPut]
        public ActionResult Delivery(int id)
        {
            string jwt = Request.Cookies["jwt"];
            if (!string.IsNullOrEmpty(jwt))
            {
                var role = getCurrentRole(jwt);
                if (role == "4"|| role == "1")
                {
                    int res = _service.Delivery(id);
                    if (res > 0) return Ok();
                    else return BadRequest();
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
