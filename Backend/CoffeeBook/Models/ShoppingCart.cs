using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class ShoppingCart
    {
        [Key]
        private int id;
        private int userId;
        [ForeignKey("userId")]
        private User user;

        private ICollection<ShoppingCart_Product> shoppingCart_Products;

        public int Id { get => id; set => id = value; }
        public int UserId { get => userId; set => userId = value; }
        public User User { get => user; set => user = value; }
        public ICollection<ShoppingCart_Product> ShoppingCart_Products { get => shoppingCart_Products; set => shoppingCart_Products = value; }
    }
}
