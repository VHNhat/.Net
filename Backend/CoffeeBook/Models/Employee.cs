using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class Employee
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        private int id;

        [Column(TypeName = "nvarchar(255)")]
        [Required]
        private string name;

        [Required]
        private int age;

        [Required]
        private bool gender;

        [Column(TypeName = "nvarchar(10000)")]
        [Required]
        private string address;

        private long salary;

        [Required]
        private string status;

        private int storeId;
        [ForeignKey("storeId")]
        private Store store;

        /*private int managerId;
        [ForeignKey("managerId")]*/
        /*private Manager manager;*/

        public int Id { get => id; set => id = value; }
        public string Name { get => name; set => name = value; }
        public int Age { get => age; set => age = value; }
        public bool Gender { get => gender; set => gender = value; }
        public string Address { get => address; set => address = value; }
        public long Salary { get => salary; set => salary = value; }
        public string Status { get => status; set => status = value; }
        public int StoreId { get => storeId; set => storeId = value; }
        public Store Store { get => store; set => store = value; }
        //public Manager Manager { get => manager; set => manager = value; }
        /*public int ManagerId { get => managerId; set => managerId = value; }*/
    }
}
