using CoffeeBook.DataAccess;
using CoffeeBook.Dto;
using CoffeeBook.Models;
using Microsoft.Extensions.Configuration;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Services
{
    public class BillService
    {
        private readonly IConfiguration _config;
        private readonly string sqlDataSource;
        private readonly Context ctx;

        public BillService()
        {
        }
        public BillService(IConfiguration config)
        {
            _config = config;
            sqlDataSource = _config.GetConnectionString("CoffeeBook");
        }
        public BillService(IConfiguration config, Context context)
        {
            _config = config;
            sqlDataSource = _config.GetConnectionString("CoffeeBook");
            ctx = context;
        }

        public List<Bill> GetAll()
        {
            return ctx.Bills.ToList();
        }

        public int Add(Bill bill)
        {
            try
            {
                Bill newBill = new Bill();
                newBill.Address = bill.Address;
                newBill.Name = bill.Name;
                newBill.Note = bill.Note;
                newBill.Validated = bill.Validated;
                newBill.Status = bill.Status;
                newBill.PayBy = bill.PayBy;
                newBill.Phone = bill.Phone;
                newBill.Time = bill.Time;
                newBill.TotalPrice = bill.TotalPrice;
                newBill.CreatedDate = DateTime.Now;
                newBill.CustomerId = bill.CustomerId;
                ctx.Bills.Add(newBill);
                return ctx.SaveChanges();
            }
            catch
            {
                return -1;
            }
        }

        public DataTable GetSale()
        {
            string query = @"select Month(CreatedDate) as 'Month', sum(TotalPrice) as 'Sales'
                             from Bill b
                             group by CreatedDate
                             order by Month(CreatedDate) asc";

            DataTable table = new DataTable();
            string sqlDataSource = _config.GetConnectionString("CoffeeBook");
            MySqlDataReader myreader;
            using (MySqlConnection myCon = new MySqlConnection(sqlDataSource))
            {
                myCon.Open();
                using (MySqlCommand myCommand = new MySqlCommand(query, myCon))
                {
                    myreader = myCommand.ExecuteReader();
                    table.Load(myreader);

                    myreader.Close();
                    myCon.Close();
                }
            }
            return table;
        }

        public DataTable GetRevenueByDay(DateTime date)
        {
            string strDate = date.ToString("yyyy-MM-dd");
            var query = @"SELECT Id, CreatedDate, TotalPrice, PayBy
                          FROM Bill 
                          WHERE DATE(CreatedDate) = @Date
                          ORDER BY TIME(CreatedDate);";
            var table = new DataTable();
            string sqlDataSource = _config.GetConnectionString("CoffeeBook");
            using (MySqlConnection connection = new MySqlConnection(sqlDataSource))
            {
                MySqlCommand command = new MySqlCommand(query, connection);

                command.CommandText = query;
                command.Parameters.AddWithValue("Date", strDate);

                connection.Open();
                var reader = command.ExecuteReader(CommandBehavior.CloseConnection);
                if (reader.HasRows)
                    table.Load(reader);

            }
            return table;
        }

        public DataTable GetRevenueByMonth(int month, int year)
        {
            var query = @"SELECT DAY(CreatedDate), COUNT(*), SUM(TotalPrice)
                          FROM Bill
                          WHERE MONTH(CreatedDate) = @Month AND YEAR(CreatedDate) = @Year
                          GROUP BY DAY(CreatedDate)
                          ORDER BY DAY(CreatedDate);";
            var table = new DataTable();
            string sqlDataSource = _config.GetConnectionString("CoffeeBook");
            using (MySqlConnection connection = new MySqlConnection(sqlDataSource))
            {
                MySqlCommand command = new MySqlCommand(query, connection);

                command.CommandText = query;
                command.Parameters.AddWithValue("Month", month);
                command.Parameters.AddWithValue("Year", year);

                connection.Open();
                var reader = command.ExecuteReader(CommandBehavior.CloseConnection);
                if (reader.HasRows)
                    table.Load(reader);

            }
            return table;
        }

        public DataTable GetRevenueByYear(int year)
        {
            string query = @"SELECT MONTH(CreatedDate) AS Month, COUNT(*) AS SLHD, SUM(TotalPrice) AS Revenue 
                             FROM Bill
                             WHERE YEAR(CreatedDate) = @Year
                             GROUP BY MONTH(CreatedDate)
                             ORDER BY MONTH(CreatedDate);";
            var table = new DataTable();
            string sqlDataSource = _config.GetConnectionString("CoffeeBook");
            using (MySqlConnection connection = new MySqlConnection(sqlDataSource))
            {
                MySqlCommand command = new MySqlCommand(query, connection);

                command.CommandText = query;
                command.Parameters.AddWithValue("Year", year);

                connection.Open();
                MySqlDataReader myreader = command.ExecuteReader(CommandBehavior.CloseConnection);
                if (myreader.HasRows)
                {
                    table.Load(myreader);
                }

            }
            return table;
        }

        public Bill GetBillId(int id)
        {
            try
            {
                var bill = ctx.Bills.Single(s => s.CustomerId == id);
                return bill;
            }
            catch
            {
                return null;
            }
        }

        public int Purchase(BillDto dto)
        {
            try
            {
                Bill bill = new Bill();
                bill.Address = dto.Address;
                bill.Name = dto.Name;
                bill.Note = dto.Note;
                bill.PayBy = dto.PayBy;
                bill.Phone = dto.Phone;
                bill.Time = dto.Time;
                bill.TotalPrice = dto.TotalPrice;
                bill.CreatedDate = DateTime.Now;
                bill.CustomerId = dto.CustomerId;
                bill.Status = "Delivering";
                ctx.Bills.Add(bill);

                var billResult = ctx.SaveChanges();
                if (billResult >= 1)
                {
                    ShoppingCart shoppingCart = new ShoppingCart();
                    shoppingCart.CustomerId = dto.CustomerId;
                    shoppingCart.CreatedDate = DateTime.Now;
                    shoppingCart.ProductQuantity = dto.ListBill.Count();

                    ctx.ShoppingCarts.Add(shoppingCart);
                    var shoppingCartsResult = ctx.SaveChanges();
                    if (shoppingCartsResult >= 1)
                    {
                        var shoppingId = ctx.ShoppingCarts.OrderByDescending(u => u.Id).FirstOrDefault().Id;

                        foreach (ShoppingCart_Product item in dto.ListBill)
                        {
                            ShoppingCart_Product checkout = new ShoppingCart_Product();
                            checkout.ProductId = item.ProductId;
                            checkout.ShoppingCartId = shoppingId;
                            checkout.TilteSize = item.TilteSize;
                            checkout.Count = item.Count;

                            ctx.ShoppingCart_Products.Add(checkout);

                        }
                        return ctx.SaveChanges();
                    }
                    return 0;
                }
                return 0;
            }
            catch
            {
                return -1;
            }
        }

        public int Delivery(int id)
        {
            try
            {
                Bill bill = ctx.Bills.Single(s => s.Id == id);
                bill.Status = "Paid";
                return ctx.SaveChanges();
            }
            catch
            {
                return -1;
            }
        }

        public int Update(int id, Bill model)
        {
            try
            {
                Bill bill = ctx.Bills.Single(s => s.Id == id);
                bill.Address = model.Address;
                bill.Name = model.Name;
                bill.Note = model.Note;
                bill.Validated = model.Validated;
                bill.Status = model.Status;
                bill.PayBy = model.PayBy;
                bill.Phone = model.Phone;
                bill.Time = model.Time;
                bill.TotalPrice = model.TotalPrice;
                bill.CustomerId = model.CustomerId;
                return ctx.SaveChanges();
            }
            catch
            {
                return -1;
            }

        }

        public int DeleteById(int id)
        {
            try
            {
                Bill bill = ctx.Bills.Single(s => s.Id == id);
                ctx.Bills.Remove(bill);
                return ctx.SaveChanges();
            }
            catch
            {
                return -1;
            }
        }
    }
}
