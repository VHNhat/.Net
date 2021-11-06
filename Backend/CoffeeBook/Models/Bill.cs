using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class Bill
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        private int id;
        private string userId;
        [Required]
        private int validated;
        [Column(TypeName = "nvarchar(100)")]
        [Required]
        private string status;
        [Required]
        private string totalPrice;
        [ForeignKey("Userid")]
        private User user;

        public int Id { get => id; set => id = value; }
        public string UserId { get => userId; set => userId = value; }
        public int Validated { get => validated; set => validated = value; }
        public string Status { get => status; set => status = value; }
        public string TotalPrice { get => totalPrice; set => totalPrice = value; }
        public User User { get => user; set => user = value; }
    }
}
