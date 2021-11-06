using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class Discount
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        private int id;
        [Column(TypeName = "nvarchar(100)")]
        [Required]
        private string name;
        [Required]
        private int value;
        private int quantity;

        public int Id { get => id; set => id = value; }
        public string Name { get => name; set => name = value; }
        public int Value { get => value; set => this.value = value; }
        public int Quantity { get => quantity; set => quantity = value; }
    }
}
