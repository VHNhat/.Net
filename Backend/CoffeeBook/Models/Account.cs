using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class Account
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        private int id;
        [StringLength(100)]
        [Required]
        private string username;
        [StringLength(100, MinimumLength = 8, ErrorMessage = "Mật khẩu phải tối thiểu 8 kí tự")]
        [Required]
        private string password;
        private int roleId;
        [ForeignKey("roleId")]
        private Role role;

        public int Id { get => id; set => id = value; }
        public string Username { get => username; set => username = value; }
        public string Password { get => password; set => password = value; }
        public int RoleId { get => roleId; set => roleId = value; }
        public Role Role { get => role; set => role = value; }
    }
}
