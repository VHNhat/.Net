using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class ProductType
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        private int id;
        [Column(TypeName ="nvarchar(100)")]
        [Required]
        private string name;
        [Column(TypeName = "nvarchar(10000)")]
        private string description;
        private ICollection<Product> products;

        public int Id { get => id; set => id = value; }
        public string Name { get => name; set => name = value; }
        public string Description { get => description; set => description = value; }
        public ICollection<Product> Products { get => products; set => products = value; }
    }
}
