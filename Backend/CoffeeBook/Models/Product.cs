using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Models
{
    public class Product
    {
        [Key]
        private int id;

        [Column(TypeName = "nvarchar(100)")]
        [Required]
        private string name;

        [Column(TypeName = "nvarchar(100)")]
        private string description;

        [Required]
        private int price;

        [DisplayFormat(DataFormatString = "{0:dd-MM-yyyy}", ApplyFormatInEditMode = true)]
        [Required]
        private DateTime createdDate;
        private string photo;
        private int size;

        private int productTypeId;
        [ForeignKey("productTypeId")]
        private ProductType productType;

        private ICollection<ShoppingCart_Product> shoppingCart_Products;

        public int Id
        {
            get { return id; }
            set { id = value; }
        }
        public string Name
        {
            get { return name; }
            set { name = value; }
        }
        public string Description 
        {
            get { return description; }
            set { description = value; }
        }

        public int Price { get => price; set => price = value; }
        public DateTime CreatedDate { get => createdDate; set => createdDate = value; }
        public string Photo { get => photo; set => photo = value; }
        public int Size { get => size; set => size = value; }
        public int ProductTypeId { get => productTypeId; set => productTypeId = value; }
        public ProductType ProductType { get => productType; set => productType = value; }
        public ICollection<ShoppingCart_Product> ShoppingCart_Products { get => shoppingCart_Products; set => shoppingCart_Products = value; }
    }
}
