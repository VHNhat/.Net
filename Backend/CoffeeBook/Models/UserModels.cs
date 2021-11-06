using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class UserModels
    {   [Column("id")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int id { get; set; }
        [Column("username")]
        public string username { get; set; }
        [Column("password")]

        public string password { get; set; }
        [Column("email")]

        public string email { get; set; }
        [Column("phone")]
 
        public string phone { get; set; }
        [Column("name")]

        public string name { get; set; }
        [Column("avata")]
 
        public string avata { get; set; }
        [Column("address")]

        public string address { get; set; }
        [Column("gender")]

        public string gender { get; set; }
        [Column("ShoppingCartid")]
        public string ShoppingCartid { get; set; }
        [ForeignKey("ShoppingCartid")]
        public ShoppingCartModels ShoppingCart { get; set; }
        public ICollection<BillModels> Bills { get; set; }
    }
}
