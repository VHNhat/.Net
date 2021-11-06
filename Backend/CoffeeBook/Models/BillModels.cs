using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class BillModels
    {
        [Column("id")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int id { get; set; }
        [Column("Userid")]
        public string Userid { get; set; }
        [Column("Userusername")]

        public string Userusername { get; set; }  
        [Column("Userphone")]
        public string Userphone { get; set; }
      
        [Column("validated")]

        public string validated { get; set; }
        [Column("status")]

        public string status { get; set; }
        [Column("totalPrice")]

        public string totalPrice { get; set; }
        [ForeignKey("Userid")]
        public UserModels User { get; set; }
    }
}
