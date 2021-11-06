using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class User
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        private int id;
        [Required]
        private string username;
        [StringLength(100, MinimumLength = 8, ErrorMessage = "Mật khẩu phải tối thiểu 8 kí tự")]
        [Required]
        private string password;
        [Required]
        private string email;
        [StringLength(11, ErrorMessage = "sdt không được vượt quá 11 kí tự.")]
        [Required]
        private string phone;
        [Column(TypeName = "nvarchar(100)")]
        [Required]
        private string name;
        private string avata;
        [Column(TypeName = "nvarchar(10000)")]
        private string address;
        private int gender;

        /*private string shoppingCartId;
        [ForeignKey("shoppingCartid")]*/
        private ICollection<ShoppingCart> shoppingCarts;
        private ICollection<Bill> bills;

        public int Id { get => id; set => id = value; }
        public string Username { get => username; set => username = value; }
        public string Password { get => password; set => password = value; }
        public string Email { get => email; set => email = value; }
        public string Phone { get => phone; set => phone = value; }
        public string Name { get => name; set => name = value; }
        public string Avata { get => avata; set => avata = value; }
        public string Address { get => address; set => address = value; }
        public int Gender { get => gender; set => gender = value; }
        /*public string ShoppingCartid { get => shoppingCartId; set => shoppingCartId = value; }*/
        /*public ShoppingCart ShoppingCart { get => shoppingCart; set => shoppingCart = value; }*/
        public ICollection<Bill> Bills { get => bills; set => bills = value; }
        public ICollection<ShoppingCart> ShoppingCarts { get => shoppingCarts; set => shoppingCarts = value; }
    }
}
