using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class Supplier
    {
        [Key]
        private int id;
        [Column(TypeName = "nvarchar(100)")]
        [Required]
        private string name;
        [Column(TypeName = "nvarchar(10000)")]
        private string description;
        [Column(TypeName = "nvarchar(10000)")]
        private string address;
        [Column(TypeName = "nvarchar(100)")]
        private string city;
        [Column(TypeName = "nvarchar(100)")]
        private string country;
        [StringLength(11,ErrorMessage ="sdt không được vượt quá 11 kí tự.")]
        private string phone;
        private string url;

        public int Id { get => id; set => id = value; }
        public string Name { get => name; set => name = value; }
        public string Description { get => description; set => description = value; }
        public string Address { get => address; set => address = value; }
        public string City { get => city; set => city = value; }
        public string Country { get => country; set => country = value; }
        public string Phone { get => phone; set => phone = value; }
        public string Url { get => url; set => url = value; }
    }
}
