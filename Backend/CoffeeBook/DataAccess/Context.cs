using CoffeeBook.Models;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.DataAccess
{
    public class Context : DbContext
    {
        // Test new connection method
        /*private readonly IConfiguration _config;
        private string connectionString;
        public Context(IConfiguration config)
        {
            _config = config;
            connectionString = _config.GetConnectionString("CoffeeBook");
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            base.OnConfiguring(optionsBuilder);
            optionsBuilder.UseMySQL(connectionString);
        }*/
        public Context(DbContextOptions<Context> options): base(options)
        {

        }

        public DbSet<Product> Products { get; set; }
        public DbSet<Bill> Bills { get; set; }
        public DbSet<Discount> Discounts { get; set; }
        public DbSet<News> News { get; set; }
        public DbSet<ProductType> ProductTypes { get; set; }
        public DbSet<ShoppingCart> ShoppingCarts { get; set; }
        public DbSet<ShoppingCart_Product> ShoppingCart_Products { get; set; }
        public DbSet<Supplier> Suppliers { get; set; }
        public DbSet<User> Users { get; set; }

        public DbSet<Account> Accounts { get; set; }
        public DbSet<Role> Roles { get; set; }
        public DbSet<Employee> Employees { get; set; }
        public DbSet<Manager> Managers { get; set; }
        public DbSet<Store> Stores { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Product>().ToTable("Product").HasKey(c => c.Id);
            modelBuilder.Entity<Bill>().ToTable("Bill").HasKey(c => c.Id);
            modelBuilder.Entity<Discount>().ToTable("Discount").HasKey(c => c.Id);
            modelBuilder.Entity<News>().ToTable("News").HasKey(c => c.Id);
            modelBuilder.Entity<ProductType>().ToTable("ProductType").HasKey(c => c.Id);
            modelBuilder.Entity<ShoppingCart>().ToTable("ShoppingCart").HasKey(c => c.Id);
            /*modelBuilder.Entity<User>()
                .HasOne(a => a.ShoppingCart).WithOne(b => b.User)
                .HasForeignKey<ShoppingCart>(e => e.UserId);*/
            modelBuilder.Entity<ShoppingCart_Product>().ToTable("ShoppingCart_Product").HasKey(c => new { c.ShoppingCartId, c.ProductId });
            modelBuilder.Entity<Supplier>().ToTable("Supplier").HasKey(c => c.Id);
            modelBuilder.Entity<User>().ToTable("User").HasKey(c => new { c.Id, c.Username });
            modelBuilder.Entity<Product>().ToTable("Product").HasKey(c => c.Id);

            modelBuilder.Entity<Account>().ToTable("Account").HasKey(c => c.Id);
            modelBuilder.Entity<Role>().ToTable("Role").HasKey(c => c.Id);
            modelBuilder.Entity<Employee>().ToTable("Employee").HasKey(c => c.Id);
            modelBuilder.Entity<Manager>().ToTable("Manager").HasKey(c => c.Id);
            modelBuilder.Entity<Store>().ToTable("Store").HasKey(c => c.Id);
            modelBuilder.Entity<Store>().HasOne(a => a.Manager).WithOne(a => a.Store)
                .HasForeignKey<Manager>(c => c.StoreId);
        }
    }
}
