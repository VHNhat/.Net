using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class ShoppingCartModels
    {   [Column("id")]
        public int id { get; set; }
        [Column("userid")]
        public string userid { get; set; }
        [Column("shoppingCartProductid")]
        public string shoppingCartProductid { get; set; }
        [ForeignKey("userid")]
        public UserModels User { get; set; }
        [ForeignKey("shoppingCartProductid")]
        public ShoppingCart_ProductModels shoppingCartProduct { get; set; }
    }
}
