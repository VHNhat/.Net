using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class Store
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        private int id;

        [Required]
        private string storeName;

        private string decription;

        [Column(TypeName = "nvarchar(10000)")]
        [Required]
        private string address;

        [Column(TypeName = "nvarchar(150)")]
        [Required]
        private string country;

        [StringLength(11, ErrorMessage = "Số điện thoại không được quá 11 số")]
        [Required]
        private string phone;

        //private int managerId;
        //[ForeignKey("managerId")]*/
        private Manager manager;

        private ICollection<Employee> employees;

        public int Id { get => id; set => id = value; }
        public string StoreName { get => storeName; set => storeName = value; }
        public string Decription { get => decription; set => decription = value; }
        public string Address { get => address; set => address = value; }
        public string Country { get => country; set => country = value; }
        public string Phone { get => phone; set => phone = value; }
        public Manager Manager { get => manager; set => manager = value; }
        public ICollection<Employee> Employees { get => employees; set => employees = value; }
        //public int ManagerId { get => managerId; set => managerId = value; }
    }
}
