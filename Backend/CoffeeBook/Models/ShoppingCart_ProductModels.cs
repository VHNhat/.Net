using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class ShoppingCart_ProductModels
    {   [Column("ShoppingCartid")]
        public int ShoppingCartid { get; set; }
        [Column("Productid")]
        public int Productid{ get; set; }
    }
}
