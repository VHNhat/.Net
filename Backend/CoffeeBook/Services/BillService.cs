﻿using CoffeeBook.DataAccess;
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

        public IQueryable findAll()
        {
            /*DataTable table = new DataTable();*/

            var query = from b in ctx.Bills
                        select b;

            /*table.Load((IDataReader)query);*/

            return query;
        }

        public DataTable save(Bill bill)
        {
            DataTable table = new DataTable();
            string query = $"insert into Bill(CustomerId, Validated, Status, TotalPrice) " +
                           $"values('{bill.CustomerId}'," +
                           $"{bill.Validated}," +
                           $"'{bill.Status}'," +
                           $"'{bill.TotalPrice}')";

            MySqlDataReader myReader;
            using (MySqlConnection myCon = new MySqlConnection(sqlDataSource))
            {
                myCon.Open();
                using (MySqlCommand myCommand = new MySqlCommand(query, myCon))
                {
                    myReader = myCommand.ExecuteReader();
                    table.Load(myReader);

                    myReader.Close();
                    myCon.Close();
                }
            }
            return table;
        }

        public int Purchase(BillDto dto)
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
            bill.CustomerId = 1;
            ctx.Bills.Add(bill);

            var billResult = ctx.SaveChanges();
            if (billResult >= 1)
            {
                ShoppingCart shoppingCart = new ShoppingCart();
                shoppingCart.CustomerId = 1;
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

        public DataTable update(Bill bill)
        {
            DataTable table = new DataTable();
            string query = @$"update Bill set
                              Validated = '{bill.Validated}',
                              Status = '{bill.Status}',
                              TotalPrice = '{bill.TotalPrice}'
                              where id = {bill.Id}";

            MySqlDataReader myReader;
            using (MySqlConnection myCon = new MySqlConnection(sqlDataSource))
            {
                myCon.Open();
                using (MySqlCommand myCommand = new MySqlCommand(query, myCon))
                {
                    myReader = myCommand.ExecuteReader();
                    table.Load(myReader);

                    myReader.Close();
                    myCon.Close();
                }
            }
            return table;
        }

        public DataTable DeleteById(int id)
        {
            DataTable table = new DataTable();
            string query = @$"delete from Bill 
                              where id = {id}";

            MySqlDataReader myReader;
            using (MySqlConnection myCon = new MySqlConnection(sqlDataSource))
            {
                myCon.Open();
                using (MySqlCommand myCommand = new MySqlCommand(query, myCon))
                {
                    myReader = myCommand.ExecuteReader();
                    table.Load(myReader);

                    myReader.Close();
                    myCon.Close();
                }
            }
            return table;
        }
    }
}
