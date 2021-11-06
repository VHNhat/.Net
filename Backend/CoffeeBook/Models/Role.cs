using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class Role
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        private int id;

        [Column(TypeName = "nvarchar(100)")]
        [MaxLength(100, ErrorMessage = "Role name không được quá 100 ký tự")]
        private string roleName;

        [Column(TypeName = "nvarchar(10000)")]
        [MaxLength(500)]
        private string description;

        private ICollection<Account> accounts;

        public int Id { get => id; set => id = value; }
        public string RoleName { get => roleName; set => roleName = value; }
        public string Description { get => description; set => description = value; }
        public ICollection<Account> Accounts { get => accounts; set => accounts = value; }
    }
}
