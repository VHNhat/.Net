using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class ShoppingCart_Product
    {
        [Key]
        private int shoppingCartId;
        [Key]
        private int productId;

        [ForeignKey("shoppingCartId")]
        private ShoppingCart shoppingCart;
        [ForeignKey("productId")]
        private Product product;

        public int ShoppingCartId { get => shoppingCartId; set => shoppingCartId = value; }
        public int ProductId { get => productId; set => productId = value; }
        public ShoppingCart ShoppingCart { get => shoppingCart; set => shoppingCart = value; }
        public Product Product { get => product; set => product = value; }
    }
}
