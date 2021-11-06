using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class CoffeeContext:DbContext
    {
        public CoffeeContext(DbContextOptions<CoffeeContext> options) : base(options) { }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<BillModels>().ToTable("Bill").HasKey(m => m.id);
            modelBuilder.Entity<ShoppingCart_ProductModels>().ToTable("ShoppingCart_Product").HasKey(m => new {m.Productid,m.ShoppingCartid });
            modelBuilder.Entity<UserModels>().ToTable("User").HasKey(m => m.id);
            modelBuilder.Entity<ShoppingCartModels>().ToTable("ShoppingCart").HasKey(m => m.id);
        }
        public DbSet<UserModels> Users { get; set; }
        public DbSet<BillModels> Bills { get; set; }
        public DbSet<ShoppingCart_ProductModels> ShoppingCart_Products { get; set; }
        public DbSet<ShoppingCartModels> ShoppingCarts { get; set; }
    }
}
