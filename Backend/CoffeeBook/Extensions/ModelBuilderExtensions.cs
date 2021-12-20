using CoffeeBook.Models;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Extensions
{
    public static class ModelBuilderExtensions
    {
        public static void Seed(this ModelBuilder modelBuilder)
        {
            #region Account
            modelBuilder.Entity<Account>().HasData(
                new Account()
                {
                    Id = 1,
                    Username = "admin",
                    Password = "admin123",
                    Name = "Võ Hoàng Nhật",
                    Avatar = "",
                    RoleId = 1
                },
                new Account()
                {
                    Id = 2,
                    Username = "manager",
                    Password = "manager123",
                    Name = "Bùi Quốc Huy",
                    Avatar = "",
                    RoleId = 2
                },
                new Account()
                {
                    Id = 3,
                    Username = "staff",
                    Password = "staff123",
                    Name = "Nguyễn Văn Nhật Huy",
                    Avatar = "",
                    RoleId = 3
                },
                new Account()
                {
                    Id = 4,
                    Username = "shipper",
                    Password = "shipper123",
                    Name = "Nguyễn Bá Hoàng",
                    Avatar = "",
                    RoleId = 4
                }
                );
            #endregion
            #region Bill
            modelBuilder.Entity<Bill>().HasData(
                new Bill()
                {
                    Id = 1,
                    Status = "Đã Thanh Toán",
                    TotalPrice = 10000000,
                    Validated = 1,
                    CustomerId = 1,
                    Address = "Đặng Thùy Trâm",
                    Name = "Võ Hoàng Nhật",
                    CreatedDate = DateTime.Parse("1/10/2021"),
                    Note = "Để đá riêng",
                    PayBy = "Tiền mặt",
                    Phone = "0942400722",
                    Time = "15-20 phút",
                },
                new Bill()
                {
                    Id = 2,
                    Status = "Đã Thanh Toán",
                    TotalPrice = 12000000,
                    Validated = 1,
                    CustomerId = 1,
                    Address = "Đặng Thùy Trâm",
                    Name = "Võ Hoàng Nhật",
                    CreatedDate = DateTime.Parse("2/10/2021"),
                    Note = "Để đá riêng",
                    PayBy = "Tiền mặt",
                    Phone = "0942400722",
                    Time = "15-20 phút",
                },
                new Bill()
                {
                    Id = 3,
                    Status = "Đã Thanh Toán",
                    TotalPrice = 5000000,
                    Validated = 1,
                    CustomerId = 1,
                    Address = "Đặng Thùy Trâm",
                    Name = "Võ Hoàng Nhật",
                    CreatedDate = DateTime.Parse("3/10/2021"),
                    Note = "Để đá riêng",
                    PayBy = "Tiền mặt",
                    Phone = "0942400722",
                    Time = "15-20 phút",
                },
                new Bill()
                {
                    Id = 4,
                    Status = "Đã Thanh Toán",
                    TotalPrice = 8000000,
                    Validated = 1,
                    CustomerId = 1,
                    Address = "Đặng Thùy Trâm",
                    Name = "Võ Hoàng Nhật",
                    CreatedDate = DateTime.Parse("4/10/2021"),
                    Note = "Để đá riêng",
                    PayBy = "Tiền mặt",
                    Phone = "0942400722",
                    Time = "15-20 phút",
                },
                new Bill()
                {
                    Id = 5,
                    Status = "Đã Thanh Toán",
                    TotalPrice = 8000000,
                    Validated = 1,
                    CustomerId = 1,
                    Address = "Đặng Thùy Trâm",
                    Name = "Võ Hoàng Nhật",
                    CreatedDate = DateTime.Parse("5/10/2021"),
                    Note = "Để đá riêng",
                    PayBy = "Tiền mặt",
                    Phone = "0942400722",
                    Time = "15-20 phút",
                },
                new Bill()
                {
                    Id = 6,
                    Status = "Đã Thanh Toán",
                    TotalPrice = 8000000,
                    Validated = 1,
                    CustomerId = 1,
                    Address = "Đặng Thùy Trâm",
                    Name = "Võ Hoàng Nhật",
                    CreatedDate = DateTime.Parse("6/10/2021"),
                    Note = "Để đá riêng",
                    PayBy = "Tiền mặt",
                    Phone = "0942400722",
                    Time = "15-20 phút",
                },
                new Bill()
                {
                    Id = 7,
                    Status = "Đã Thanh Toán",
                    TotalPrice = 9000000,
                    Validated = 1,
                    CustomerId = 1,
                    Address = "Đặng Thùy Trâm",
                    Name = "Võ Hoàng Nhật",
                    CreatedDate = DateTime.Parse("7/10/2021"),
                    Note = "Để đá riêng",
                    PayBy = "Tiền mặt",
                    Phone = "0942400722",
                    Time = "15-20 phút",
                },
                new Bill()
                {
                    Id = 8,
                    Status = "Đã Thanh Toán",
                    TotalPrice = 13000000,
                    Validated = 1,
                    CustomerId = 1,
                    Address = "Đặng Thùy Trâm",
                    Name = "Võ Hoàng Nhật",
                    CreatedDate = DateTime.Parse("8/10/2021"),
                    Note = "Để đá riêng",
                    PayBy = "Tiền mặt",
                    Phone = "0942400722",
                    Time = "15-20 phút",
                },
                new Bill()
                {
                    Id = 9,
                    Status = "Đã Thanh Toán",
                    TotalPrice = 15000000,
                    Validated = 1,
                    CustomerId = 1,
                    Address = "Đặng Thùy Trâm",
                    Name = "Võ Hoàng Nhật",
                    CreatedDate = DateTime.Parse("9/10/2021"),
                    Note = "Để đá riêng",
                    PayBy = "Tiền mặt",
                    Phone = "0942400722",
                    Time = "15-20 phút",
                },
                new Bill()
                {
                    Id = 10,
                    Status = "Đã Thanh Toán",
                    TotalPrice = 17000000,
                    Validated = 1,
                    CustomerId = 1,
                    Address = "Đặng Thùy Trâm",
                    Name = "Võ Hoàng Nhật",
                    CreatedDate = DateTime.Parse("10/10/2021"),
                    Note = "Để đá riêng",
                    PayBy = "Tiền mặt",
                    Phone = "0942400722",
                    Time = "15-20 phút",
                },
                new Bill()
                {
                    Id = 11,
                    Status = "Đã Thanh Toán",
                    TotalPrice = 16000000,
                    Validated = 1,
                    CustomerId = 1,
                    Address = "Đặng Thùy Trâm",
                    Name = "Võ Hoàng Nhật",
                    CreatedDate = DateTime.Parse("11/10/2021"),
                    Note = "Để đá riêng",
                    PayBy = "Tiền mặt",
                    Phone = "0942400722",
                    Time = "15-20 phút",
                }
                );
            #endregion
            #region Customer
            modelBuilder.Entity<Customer>().HasData(
                new Customer()
                {
                    Id = 1,
                    Username = "khachhang1",
                    Password = "khachhang1",
                    Email = "vohoangnhat06052001@gmail.com",
                    Name = "Võ Hoàng Nhật",
                    Gender = 1,
                    Address = "167/19 Đặng Thùy Trâm",
                    Avata = "",
                    City = "Hồ Chí Minh",
                    Country = "Việt Nam",
                    Phone = "0942400722",
                },
                new Customer()
                {
                    Id = 2,
                    Username = "khachhang2",
                    Password = "khachhang2",
                    Email = "khachhang2@gmail.com",
                    Name = "Bùi Quốc Huy",
                    Gender = 1,
                    Address = "167/19 Đặng Thùy Trâm",
                    Avata = "",
                    City = "Hồ Chí Minh",
                    Country = "Việt Nam",
                    Phone = "0942400723",
                },
                new Customer()
                {
                    Id = 3,
                    Username = "khachhang3",
                    Password = "khachhang3",
                    Email = "khachhang3@gmail.com",
                    Name = "Nguyễn Bá Hoàng",
                    Gender = 1,
                    Address = "167/19 Đặng Thùy Trâm",
                    Avata = "",
                    City = "Hồ Chí Minh",
                    Country = "Việt Nam",
                    Phone = "0942400724",
                },
                new Customer()
                {
                    Id = 4,
                    Username = "khachhang4",
                    Password = "khachhang4",
                    Email = "khachhang4@gmail.com",
                    Name = "Nguyễn Văn Nhật Huy",
                    Gender = 1,
                    Address = "167/19 Đặng Thùy Trâm",
                    Avata = "",
                    City = "Hồ Chí Minh",
                    Country = "Việt Nam",
                    Phone = "0942400725",
                }
                );
            #endregion
            #region Discount
            modelBuilder.Entity<Discount>().HasData(
                new Discount()
                {
                    Id = 1,
                    Name = "Phiếu giảm giá 20k cho đơn hàng trên 100k",
                    Quantity = 10,
                    Value = 20000,
                    Photo = "https://minio.thecoffeehouse.com/image/admin/storage/852__C4_90o_CC_82_CC_80ng-gia_CC_81-29K_coupon.jpg",
                    ExpiredDate = DateTime.Parse("1/5/2022"),
                    MinPrice = 100000
                },
                new Discount()
                {
                    Id = 2,
                    Name = "Phiếu giảm giá 30k cho đơn hàng trên 150k",
                    Quantity = 10,
                    Value = 30000,
                    Photo = "https://minio.thecoffeehouse.com/image/admin/Coupondelivery30_684527.jpg",
                    ExpiredDate = DateTime.Parse("1/5/2022"),
                    MinPrice = 150000
                },
                new Discount()
                {
                    Id = 3,
                    Name = "Phiếu giảm giá 50k cho đơn hàng trên 300k",
                    Quantity = 10,
                    Value = 50000,
                    Photo = "https://minio.thecoffeehouse.com/image/admin/Coupondelivery30_684527.jpg",
                    ExpiredDate = DateTime.Parse("1/5/2022"),
                    MinPrice = 300000
                }
                );
            #endregion
            #region Employee
            modelBuilder.Entity<Employee>().HasData(
                new Employee
                {
                    Id = 1,
                    Name = "Võ Hoàng Nhật",
                    Age = 20,
                    Gender = 1,
                    Email = "nhatvh@gmail.com",
                    Phone = "1234567890",
                    Address = "Quận 1",
                    City = "Thành phồ Hồ Chí Minh",
                    Country = "Việt Nam",
                    Salary = 100_000_000,
                    Status = "Hoạt động"
                },
                new Employee
                {
                    Id = 2,
                    Name = "Bùi Quốc Huy",
                    Age = 20,
                    Gender = 1,
                    Email = "huybq@gmail.com",
                    Phone = "1234567891",
                    Address = "Quận 1",
                    City = "Thành phồ Hồ Chí Minh",
                    Country = "Việt Nam",
                    Salary = 100_000_000,
                    Status = "Hoạt động"
                },
                new Employee
                {
                    Id = 3,
                    Name = "Nguyễn Văn Nhật Huy",
                    Age = 20,
                    Gender = 1,
                    Email = "huynvn@gmail.com",
                    Phone = "1234567892",
                    Address = "Quận 1",
                    City = "Thành phồ Hồ Chí Minh",
                    Country = "Việt Nam",
                    Salary = 100_000_000,
                    Status = "Hoạt động"
                },
                new Employee
                {
                    Id = 4,
                    Name = "Nguyễn Bá Hoàng",
                    Age = 20,
                    Gender = 1,
                    Email = "hoangnb@gmail.com",
                    Phone = "1234567893",
                    Address = "Quận 1",
                    City = "Thành phồ Hồ Chí Minh",
                    Country = "Việt Nam",
                    Salary = 100_000_000,
                    Status = "Hoạt động"
                }
                );
            #endregion
            #region Manager 
            modelBuilder.Entity<Manager>().HasData(
                new Manager
                {
                    Id = 1,
                    Name = "Võ Hoàng Nhật",
                    Age = 20,
                    Gender = 1,
                    Email = "nhatvh@gmail.com",
                    Phone = "1234567890",
                    Address = "Quận 1",
                    City = "Thành phồ Hồ Chí Minh",
                    Country = "Việt Nam",
                    Salary = 100_000_000,
                    Status = "Hoạt động"
                },
                new Manager
                {
                    Id = 2,
                    Name = "Bùi Quốc Huy",
                    Age = 20,
                    Gender = 1,
                    Email = "huybq@gmail.com",
                    Phone = "1234567891",
                    Address = "Quận 1",
                    City = "Thành phồ Hồ Chí Minh",
                    Country = "Việt Nam",
                    Salary = 100_000_000,
                    Status = "Hoạt động"
                },
                new Manager
                {
                    Id = 3,
                    Name = "Nguyễn Văn Nhật Huy",
                    Age = 20,
                    Gender = 1,
                    Email = "huynvn@gmail.com",
                    Phone = "1234567892",
                    Address = "Quận 1",
                    City = "Thành phồ Hồ Chí Minh",
                    Country = "Việt Nam",
                    Salary = 100_000_000,
                    Status = "Hoạt động"
                },
                new Manager
                {
                    Id = 4,
                    Name = "Nguyễn Bá Hoàng",
                    Age = 20,
                    Gender = 1,
                    Email = "hoangnb@gmail.com",
                    Phone = "1234567893",
                    Address = "Quận 1",
                    City = "Thành phồ Hồ Chí Minh",
                    Country = "Việt Nam",
                    Salary = 100_000_000,
                    Status = "Khóa"
                }
                );
            #endregion
            #region News
            modelBuilder.Entity<News>().HasData(
                new News()
                {
                    Id = 1,
                    Thumbnail = "https://file.hstatic.net/1000075078/file/1_f225c98d69ec4dfe8bfb5df29a25e13e_grande.jpg",
                    Title = "CÂU CHUYỆN VỀ LỊCH SỬ VÀ NGUỒN GỐC XUẤT XỨ CỦA CÀ PHÊ",
                    Content = @"Truyền thuyết về quả cà phê
Có rất nhiều truyền thuyết về nguồn gốc của cà phê.Tuy nhiên,câu chuyện người đàn ông chăn dê tên Kaldi phát hiện ra cây cà phê ở khu rừng cổ của cao nguyên Ethiopia là phổ biến và được nhiều người tin nhất.
Kaldi đã khám phá ra cà phê sau khi những con dê của anh ta ăn một loại quả màu đỏ từ cái cây có những bông hoa trắng và chúng trở nên hăng hái đến mức không đi ngủ vào ban đêm. Anh ta cũng đã ăn thử và nó giúp anh tỉnh táo trong suốt nhiều giờ cầu nguyện vào buổi tối. Sau đó, Kaldi đã báo cáo phát hiện của mình với giáo sĩ của tu viện địa phương. Vị giáo sĩ chia sẻ khám phá của Kaldi với nhiều người khác và kết quả là thông tin về một loại quả tràn đầy năng lượng nhanh chóng được lan truyền.
Lịch sử và hành trình du nhập của cà phê
Tuy nhiên, không chỉ có truyền thuyết, những ghi chép và dấu tích còn lại cũng khiến người ta tin rằng Kaffa - Ethiopia chính là vùng đất khởi nguyên của cây cà phê từ thế kỷ thứ 9. Và đến thế kỷ 14, những người buôn nô lệ đã mang cà phê từ Ethiopia sang xứ Ả Rập. Cà phê trở thành thức uống phổ biến của người Ả Rập với cách chế biến đầu tiên chỉ đơn giản là tách lấy hạt và cho vào nấu trong nước sôi. Ả Rập đã trở thành nơi trồng cà phê độc quyền với trung tâm giao dịch cà phê là thành phố cảng Mocha, hay còn được gọi là Mokka, tức là thành phố Al Mukha thuộc Yemen ngày nay. 
Người Ả Rập rất tự hào và giữ bí mật để bảo tồn độc quyền loại thức uống được chế biến từ hạt cà phê. Họ đưa ra những quy định rất chặt chẽ để gìn giữ sự độc quyền về cà phê của mình như chỉ được phép mang hạt cà phê đã rang ra khỏi xứ, cũng như không cho người ngoại quốc bén mảng đến các đồn điền cà phê. Thế nhưng, dù nghiêm ngặt đến mức nào thì cũng có người vượt qua được, những khách hành hương được thưởng thức nước cà phê đã lén lút mang hạt giống về trồng, chẳng bao lâu khắp khu vực Trung Đông đều có trồng và truyền đi mỗi lúc một xa hơn. Bước sang thế kỷ 15 thì cà phê đến Armenia, Persia, Thổ Nhĩ Kỳ và phía bắc Châu Phi, phổ biến trong thế giới Hồi giáo và lan tỏa đến Ý, sau đó là phần còn lại của Châu Âu, Indonesia, Mĩ. 
Bắt nguồn tên gọi “Cà phê”
Nhiều nhà nghiên cứu cho rằng, chữ “Cà phê” có thể bắt nguồn từ chữ “Qahwa” (theo tiếng Ả Rập thì “Qahwa” có nghĩa là thức uống được tạo ra từ nước của hạt quả). Qua đến người Thổ Nhĩ Kỳ được đọc là “Kahve” và nó đã được sử dụng rộng rãi khắp nơi. Một số người khác thì lại cho rằng, chữ “Cà phê” được đặt theo tên của thị trấn “Kaffa” của Ethiopia. Tuy nhiên, một số khác lại nói thị trấn “Kaffa” được đặt theo tên “Cà phê”. Và từ “cà phê” mà Việt Nam ta vẫn hay dùng có gốc từ chữ Café trong tiếng Pháp, một loại thức uống màu đen có chứa chất caffeine.
Như vậy, có thể thấy cà phê đã xuất hiện từ rất lâu đời và có những thông tin còn hoài nghi chưa được xác thực. Thế nhưng, không còn quá quan trọng vì nó đã lan khắp thế giới và ngày càng hiện diện, hoà nhịp trong cuộc sống của tất cả chúng ta. Mỗi ngày, thế giới ước tính có 2,25 tỷ tách cà phê được bán ra và đối với nhiều người, cà phê đã trở thành một phần không thể thiếu. 
Đặc biệt, không thể phủ nhận những giá trị vô hình mà cà phê đang mang lại. Ngoài là một thức uống thú vị và yêu thích của rất nhiều người, nó là một sự khởi đầu của ngày mới, khởi đầu của những câu chuyện, sự gắn kết của những mối quan hệ, đưa chúng ta gần nhau hơn, sẻ chia và kết nối tình cảm. Và đó cũng chính là những giá trị mà The Coffee House luôn hướng đến và đem lại bên cạnh những tách cà phê tinh túy và chất lượng, những ly trà đượm hương cùng cảm hứng về lối sống hiện đại, năng động mà gần gũi. "
                },
                new News()
                {
                    Id = 2,
                    Thumbnail = "https://file.hstatic.net/1000075078/file/1_3f3c6d40508d4a6aba9fa9e5aad7708a_grande.jpg",
                    Title = "THƯỞNG THỨC CÀ PHÊ THEO PHONG CÁCH MỸ",
                    Content = @"Ở Mỹ, cà phê được xuất hiện từ những năm 1600 tại New Amsterdam - thành phố New York nổi tiếng ngày nay. Từ đó cà phê dần được sử dụng và gần gũi hơn. Tuy nhiên, phong cách cà phê hiện đại của Mỹ như hiện nay được cho rằng bắt nguồn từ Ý và phát triển mạnh mẽ trong khoảng 20 năm trở lại đây.
Dù chịu ảnh hưởng bởi văn hóa cà phê du nhập từ Ý với lối pha dùng máy Barista nhanh chóng phổ biến,
                    tuy nhiên lại có sự khác biệt rõ rệt giữa hai nền văn hoá cà phê của Ý và Mỹ.Đến với Ý bạn sẽ thấy sự chờ đợi một cách chậm rãi,
                    nhẹ nhàng và sành điệu.Còn với người Mỹ thì luôn trong tâm thái nhanh gọn,
                    tranh thủ thời gian.Điều này hoàn toàn dễ hiểu khi Mỹ là một quốc gia làm việc với cường độ cao,
                    con người rất quý trọng thời gian và yêu cầu sự gọn lẹ. 
Là quốc gia trẻ và đầy sức sống, con người ở Mỹ chuộng lối sống tự do, tất cả đều theo sở thích cá nhân và văn hóa cà phê cũng không ngoại lệ. Người Mỹ sử dụng cà phê hoàn toàn theo ý thích, không sành điệu như người châu Âu, cũng không cầu kì như người Ả Rập, mà đề cao tính thoải mái.
Cà phê Việt thì luôn đậm đà, còn cà phê của người Mỹ được cho là nhạt từ màu sắc đến cả hương vị. Cà phê Mỹ là chất nước loãng màu nâu nhạt, có vị hơi khét, vậy nên người ta thường bảo cho thêm hai, thậm chí ba gói đường và nửa ly sữa cũng không làm chất nước ấy ngọt và thơm hơn.
Người Mỹ uống cà phê mọi lúc mọi nơi và nhiều lần trong ngày. Cà phê đối với họ là thức uống giúp tập trung đầu óc, tỉnh táo, chống buồn ngủ. Món Americano, một thức uống được biến tấu nhẹ nhàng từ món Espresso của người Ý được xem là thức uống truyền thống của người Mỹ. Ngoài ra, ở đất nước này phổ biến với Iced Coffee – cà phê đặc, nóng, thêm đường được rót vào một ly đựng đá. 
Vào quán cà phê, người Mỹ không chỉ đơn thuần là tán gẫu. Đôi lúc họ làm việc trong quán cà phê. Đây vừa là cách tận dụng thời gian, hoặc thay đổi không gian làm việc để đầu óc tự do hơn là ngồi trong văn phòng nhàm chán. Hơn thế nữa, người ta còn vào cà phê để bàn bạc công việc, thảo luận, trao đổi và dĩ nhiên là có cả hẹn hò! Và những điều này thì chúng ta cũng thấy thân thuộc đối với người Việt và phong cách đi cà phê hiện nay. 
Cũng giống khi bước vào không gian của The Coffee House, bạn sẽ bắt gặp hình ảnh quen thuộc của những đôi bạn trẻ trò chuyện bên tách cà phê, ai đó làm việc hăng say hay chú tâm đọc sách bất kể những âm thanh rì rào bên cạnh, cũng như những cuộc gặp gỡ niềm nở giữa đối tác và khách hàng, hay những người bạn lâu ngày mới gặp lại nhưng rất đỗi thân tình... 
Đối với The Coffee House đó là một niềm thật hạnh phúc. Hạnh phúc khi luôn được đồng hành và gắn liền trong những câu chuyện, những lần gặp gỡ đong đầy tình cảm. Góp mặt trong hành trình sống của mỗi người và là nơi lưu giữ những kỷ niệm. The Coffee House tự hào mang đến những trải nghiệm khác biệt, thoải mái, những hương vị tinh tế và quyến rũ cho khách hàng thân yêu, đặc biệt biệt là góp phần kiến tạo nên phong cách và văn hóa thưởng thức cà phê của người Việt."
                },
                new News()
                {
                    Id = 3,
                    Thumbnail = "https://file.hstatic.net/1000075078/file/phap_477cd4bf793c44dd91589ef3cc5dfe22_grande.jpg",
                    Title = "PHONG CÁCH UỐNG CÀ PHÊ CỦA CÁC NƯỚC TRÊN THẾ GIỚI",
                    Content = @"Uống cà phê là một thú vui và đến với mỗi quốc gia, bạn sẽ thấy rằng thú vui này sẽ mang những nét khác biệt. The Coffee House chia sẻ đến bạn những phong cách uống cà phê nổi tiếng của các nước trên thế giới, cùng xem để hiểu hơn về văn hoá độc đáo này, hay kể cả tự pha chế để thưởng thức xem cảm nhận mới lạ như thế nào nhé!
Italy
Là “thủ phủ” cà phê nổi tiếng trên thế giới, thức uống này là món không thể thiếu đối với người dân Italy.Họ thường bắt đầu ngày mới với một ly cà phê và loại thường được dùng nhiều nhất chính là Espresso. 
Đây là loại cà phê có thể phục vụ nhanh, có cho bạn ngay lập tức và thường được đựng trong những chiếc ly nhỏ. Người Italy khi thưởng thức sẽ cầm ly trong tay, hít hà mùi thơm quyến rũ rồi uống trong 3-4 hơi. Cả quá trình diễn ra nhanh gọn nhưng đầy thanh lịch theo đúng phong cách người Italy.

Pháp
Ở Pháp thường dùng Café Au Lait, nghe thôi đã rất Pháp rồi đúng không? Đây là loại cà phê được pha chế hết sức đơn giản. 
Là sự pha trộn giữa cà phê espresso và sữa tươi nóng. Café Au Lait là thức uống yêu thích của người Pháp vào buổi sáng và thường được dùng cùng bánh mì Baguette. Café Au Lait và Baguette, sự kết hợp biểu tượng của ẩm thực Pháp.

Đức
Eiskaffee là món cà phê quen thuộc của người Đức. Nó đích thị dành cho những ai thích sự ngọt ngào. Eiskaffee trong tiếng Đức có nghĩa là “cà phê kem lạnh”. Vị beo béo của kem lạnh lơ lửng trong lớp cà phê đắng kết hợp cùng tầng kem whipping ngọt ngào và một ít vụn choco đã tạo nên thứ đồ uống hấp dẫn khó cưỡng. 
Vậy nên Eiskaffee đã đã chinh phục khẩu vị không chỉ của người dân nước Đức, trở thành món đồ uống “quốc dân” của đất nước này mà còn thu hút cả những khách du lịch đến nơi đây.

Bồ Đào Nha
Người Bồ Đào Nha cực kì ưa chuộng và phổ biến với món cà phê pha chanh mát lạnh Mazagran. Để thưởng thức món cà phê này hết sức đơn giản, bạn chỉ cần đổ hai tách espresso và nước cốt chanh vào một ly cao, thêm đường cho vừa uống và một ít đá viên. 
Với sự kết hợp đặc biệt từ vị chua của chanh khiến món cà phê này có vị khác hẳn so với các loại cà phê khác trên thế giới. Đây cũng là một thức uống lý tưởng và sảng khoái cho những ngày hè.

Hy Lạp
Khá giống như người Bồ Đào Nha, người Hy Lạp cũng có một thức uống cà phê được yêu chuộng vào những ngày hè nhằm xua đi cái nóng nực. Đó chính là Frappé. Trong tiếng Pháp nó có nghĩa “ướp lạnh”, là loại cà phê hòa tan đánh sủi bọt hết sức độc đáo và được biết đến như biểu tượng của đất nước Hy Lạp. 
Để pha Frappé, đầu tiên cho cà phê bột vào ly, thêm một thìa cà phê đường và một ít nước đá. Dùng máy trộn tay (handmixer) khuấy mạnh cho đến khi cà phê sủi bọt. Cuối cùng cho đá và rót một thêm một ít nước lạnh là bạn đã hoàn thành. Frappe thường được pha chế khá đậm và sử dụng cà phê hoà tan mạnh, nếu dùng không quen, bạn có thể thêm nước để pha loãng cà phê hoặc thêm đường vào để thưởng thức ngọt hơn.

Cà phê sữa Việt Nam
Mỗi đất nước đều có loại cà phê yêu thích đặc trưng. Và đối với người Việt đó là món cà phê sữa. Bạn thường thưởng thức cà phê sữa như thế nào? Mỗi buổi sáng tự pha cho mình hay ra cửa hàng gọi một ly cà phê sữa theo ý thích. Không chỉ quen thuộc với người Việt, cà phê sữa là món đồ uống mà khách du lịch các nước rất hứng thú thưởng thức khi đặt chân đến mảnh đất hình chữ S. Và nhiều thực khách nước ngoài khi vào cửa hàng của The Coffee House cũng thường xuyên lựa chọn cà phê sữa.
Không giống những nước khác thường sử dụng cà phê pha sẵn, cà phê bột,…Việt Nam sử dụng loại cà phê xay nhuyễn, cho vào loại phin đặc trưng của bản xứ, sau đó đổ nước sôi vào. Từng giọt cà phê đậm đặc, sóng sánh, chảy qua chiếc phin đặc biệt hòa cùng sữa đặc, cho thêm đá đã tạo nên ly cà phê sữa đá đậm đà, đắng đót nhưng cũng thanh ngọt nơi hậu vị.
Không chỉ mỗi quốc gia sẽ có một loại cà phê đặc trưng và cách thưởng thức riêng biệt, mà mỗi cá nhận lại còn có cả phong cách uống cà phê rất khác. Phong cách thưởng thức cà phê của bạn là gì? Cà phê đen hay cà phê sữa, nhiều đường hay ít sữa, nóng hay đá, pha phin hay dùng sẵn, lựa chọn Espresso hay Capuccino, bạn có muốn thưởng thức thêm vị Matcha hay Macchiato... Mỗi người mỗi sở thích, mỗi người mỗi phong cách và cá tính, nhưng The Coffee House lại luôn tự tin mang đến sự hài lòng cho khách hàng bởi những trải nghiệm độc đáo, sự đa dạng, phá cách và hơn hết là những hương vị tuyệt vời nhất."
                },
                new News()
                {
                    Id = 4,
                    Thumbnail = "https://file.hstatic.net/1000075078/file/1_dda8641750b14bc4aa1486fe9a01f1d0_grande.jpg",
                    Title = "GỢI Ý NHỮNG LOẠI CÀ PHÊ CHO NGÀY BẬN RỘN",
                    Content = @"GỢI Ý NHỮNG LOẠI CÀ PHÊ CHO NGÀY BẬN RỘN
Đối với những người đã yêu thích cà phê rồi thì hôm nào cũng cần phải có ít nhất một ly.Và đặc biệt đối với những người đã là tín đồ của cà phê thì họ lại rất khó tính,
                    yêu cầu cao về hương vị và sự nguyên chất,
                    đậm đà của loại cà phê mà mình thưởng thức mỗi ngày.Bạn có đang nằm trong số đó,
                    và nếu những ngày bận rộn,
                    bạn không có thời gian để tự tay pha cho mình một ly cà phê hay ghé đến quán cà phê quen thuộc để thưởng thức thì như thế nào ?

Đừng nói rằng bạn sẽ tạm hoãn cái ghiền cà phê của bản thân lại nhé.Vì đó thật sự là một điều rất khó chịu,
                    đặc biệt nữa là khi bạn đã quá quen với việc bắt đầu buổi sáng với một ly cà phê để giúp tỉnh táo,
                    tập trung và có cảm hứng hơn khi bắt đầu bước vào công việc.

The Coffee House gợi ý những loại cà phê đậm đà,
                    thơm ngon và hương vị đúng nghĩa cà phê tinh túy.Hãy cùng lựa chọn cho những ngày bận rộn trong guồng quay của công việc và cuộc sống nhé.Với những dòng cà phê được giới thiệu dưới đây,
                    The Coffee House tin chắc dù thời gian hạn hẹp đến đâu thì bạn cũng có thể ưu ái và phục vụ bản thân mình bằng một tách cà phê thơm ngon đúng “gu” nhất.

Cà phê sữa đá hoà tan

Với cà phê sữa đá hoà tan thì thật dễ dàng để bạn bắt đầu ngày mới với ly cà phê sóng sánh,
                    thơm lừng và lựa chọn uống nóng hay uống lạnh theo sở thích của bản thân.Chắc chắn ai cũng phải công nhận sự tiện lợi của cà phê sữa đá hoà tan vì không cần phải mất nhiều thời gian pha chế hay phải chuẩn bị những dụng cụ pha chế.Thay vào đó chỉ cần pha theo hướng dẫn trên bao bì là đã có ngay một ly cà phê thơm ngon,
                    chuẩn vị.The Coffee House gợi ý 3 dòng sản phẩm để bạn lựa chọn: cà phê sữa đá hoà tan dạng túi,
                    cà phê sữa đá hoà tan dạng hộp và cà phê hoà tan 3in1 đậm vị Việt. 
Cà phê sữa đá hoà tan dạng túi và hộp: Đây là 02 phiên bản giống nhau về chất lượng cà phê, giữ được hương vị thơm ngon thuần tuý như cà phê pha phin với sữa đặc. Vị đắng thanh của cà phê hoà quyện với vị ngọt béo của sữa, giúp bạn luôn tỉnh táo và hứng khởi cho ngày làm việc thật hiệu quả. Thật sự là lựa chọn lý tưởng cho những ngày hối hả của nhịp sống hiện đại. 
Cà phê hoà tan 3in1 đậm vị Việt: đây là dòng sản phẩm với vị đậm đà, đặc biệt phù hợp với những ai có gu cà phê mạnh. 3in1 đậm vị Việt giữ trọn vẹn hương vị đắng thơm tinh tế từ hạt cà phê, là sự lựa chọn đậm đà và mạnh mẽ hơn so với hai phiên bản cà phê sữa dạng hộp và túi được The Coffee House chia sẻ trên. 

Cà phê sữa đá dạng lon

Cà phê sữa đá dạng lon của The Coffee House mang phong cách hiện đại, đồng hành cùng nhịp sống sôi nổi của người trẻ và tiện dụng trong cuộc sống tất bật . Sản phẩm được sản xuất từ những hạt cà phê Việt thượng hạng và công thức phối trộn độc đáo mang đến hương vị đậm đà, thơm béo. 
Với thiết kế lon cao trẻ trung, đơn giản và tiện lợi, dù ở đâu, bạn vẫn có thể tận hưởng vị cà phê sữa thơm ngon đến ngụm cuối cùng. Sản phù hợp cả khi bạn ở nhà, văn phòng, đặc biệt là những lúc di chuyển trên đường làm việc, công tác và cả những nẻo đường đi phượt, du lịch cùng bạn bè,...  

Cà phê pha phin 

Nếu bạn là người dành tình yêu cuồng nhiệt với cà phê pha phin, khi có thời gian nhiều hơn một chút bạn có thể tạt ngang hoặc nhấc điện thoại đặt một ly cà phê sữa, cà phê đen của The Coffee House nhé. Với lợi thế cửa hàng phủ sóng rộng khắp cùng đội ngũ nhân viên pha chế, giao hàng chuyên nghiệp, The Coffee House sẽ nhanh chóng gửi đến bạn hương vị cà phê yêu thích mà không cần phải đợi chờ quá lâu. 

Như vậy, bạn vừa có thể làm việc vừa thưởng thức hương vị cà phê chuẩn gu hoặc tự bù đắp sau khi công việc đã xong xuôi bằng ly cà phê The Coffee House được pha phin truyền thống với hương vị đậm đà, hài hòa giữa vị ngọt đầu lưỡi và vị đắng thanh thoát nơi hậu vị.
Bạn biết không, theo nhiều nghiên cứu đã chứng minh rằng uống cà phê mỗi ngày với một lượng vừa đủ sẽ rất tốt cho cơ thể vì nó giảm nguy cơ mắc những bệnh nguy hiểm như: cao huyết áp, tiểu đường tuýp 2, xơ gan, parkinson, hen suyễn, sỏi mật.... Vậy nên, nếu bạn là người rất đam mê với cà phê thì vẫn nên giữ thói quen đó mỗi ngày. Thế nhưng, nên dừng ở mức 3-4 ly, vì liều lượng caffeine khuyến nghị mỗi ngày để giữ mức an toàn là 400mg, tương đương khoảng 4 ly cà phê 240ml. 

Do đó, đối với những ngày bận rộn, bạn có thể lựa chọn những loại cà phê tiện lợi và chất lượng tuyệt vời được The Coffee House gợi ý trên để vừa thỏa mãn tình yêu với cà phê, vừa đảm bảo một ngày làm việc tập trung tinh thần, hăng say và hiệu quả. Đặc biệt, giúp cơ thể nhận được những lợi ích bất ngờ từ cà phê và tránh nguy cơ mắc các bệnh nguy hiểm. "
                },
                new News()
                {
                    Id = 5,
                    Thumbnail = "https://file.hstatic.net/1000075078/file/masala_97a1863ecaa14939a38ae35d76454663_grande.jpg",
                    Title = "MASALA CHAI - HƯƠNG VỊ TRÀ SỮA ĐỘC ĐÁO CỦA NGƯỜI ẤN",
                    Content = @"Nghe đến trà sữa chắc hẳn bạn sẽ nghĩ nó là món thức uống được giới trẻ yêu thích và chỉ mới nổi lên trong những năm gần đây. Tuy nhiên, The Coffee House chia sẻ với bạn rằng hoàn toàn không phải như vậy. Bởi vì trà sữa đã có mặt từ rất lâu đời, minh chứng rõ ràng nhất chính là Trà sữa Masala Chai Ấn Độ. Đây là loại thức uống được người dân Ấn từ người già đến trẻ nhỏ đều yêu thích, nó đã có mặt hàng nghìn năm nay và ngày càng trở nên nổi tiếng trên thế giới. 

Bạn biết không, “Chai” trong tiếng Hindi của Ấn có nghĩa là “Trà”, và Masala Chai có nghĩa là “Trà hỗn hợp gia vị”. Nghe đến đây thôi cũng phần nào hình dung được sự đặc biệt của loại trà sữa này khi nó có sự kết hợp của nhiều gia vị khác nhau. Bạn có đang tò mò về những gia vị có trong Masala Chai?

Nguyên liệu chính của Masala Chai cũng giống như nhiều loại trà sữa khác đó chính là sữa tươi và trà đen. Tuy nhiên, sự độc đáo của Masala chai nằm ở chỗ kết hợp đặc biệt với những loại thảo dược có vị cay nồng nàn. Nhóm nguyên liệu chính tạo nên Trà sữa Masala Chai khác biệt với những loại trà sữa trên thị trường hiện nay đó chính là: tiêu đen, vỏ quế, đinh hương, thảo quả, bạch đậu khấu, bột gừng khô, bột nhục đậu khấu,... Đúng như tên gọi, đây quả thật là “trà hỗn hợp gia vị”. 
Để làm nên món trà sữa Masala Chai đúng chuẩn thì người Ấn sẽ đun sữa tươi trên bếp cho thật nóng, song song đó giã nát các loại nguyên liệu ra thành bột mịn rồi cho vào đun cùng với sữa. Tiếp đến, cho trà đen cùng lượng đường vừa phải vào và đun cho tan. Sau khi trà sữa sôi đều sẽ tắt bếp, đổ trà sữa vào ray lọc để loại bỏ xác trà cũng như các cặn thảo dược. Cuối cùng sẽ tạo nên một thức uống hòa quyện giữa vị ngọt ngào là vị nồng ấm, và đối với Masala Chai dù uống nóng hay lạnh đều rất tuyệt vời với những cảm nhận khó quên.

Với danh sách một loạt các thảo dược nổi bật như trên, Masala Chai không chỉ đơn giản là một thức uống giải khát, mà ở Ấn Độ đây được xem như một bài thuốc cổ truyền rất tốt cho sức khỏe. Thưởng thức một ly Trà sữa Masala Chai có thể giúp làm ấm cơ thể, lưu thông máu huyết, ngoài ra còn giúp tăng cường tiêu hoá, tăng cường trí nhớ, trị viêm khớp và các bệnh đau đầu kinh niên... Có lẽ vì chính những lý do đó mà Masala Chai tồn tại qua hàng ngàn năm, đến bây giờ càng được ưa chuộng và ngày càng được thế giới biết đến nhiều hơn. 

Dù chưa xác định được thời gian ra đời của món thức uống truyền thống đặc biệt có nguồn gốc từ Ấn Độ này (có truyền thuyết ghi lại Masala Chai đã xuất hiện 9.000 năm trước, cũng có tài liệu cho rằng nó xuất hiện cách đây khoảng 5.000 năm), tuy nhiên với bề dày và ngày càng nổi tiếng như thế, có thể khẳng định Masala Chai là một phần linh hồn của ẩm thực Ấn, một phần văn hoá và không thể thiếu trong nhịp sống của những con người nơi đây. Hơn thế nữa, Masala Chai góp phần tạo nên sự đa dạng, ấn tượng và độc đáo cho nền ẩm thực của thế giới. 
Bạn nghĩ sao nếu trong lúc tiết trời se lạnh, trong không khí rộn ràng của mùa Lễ hội cuối năm như thế này được cầm trên tay ly trà sữa Masala Chai cay nồng ấm áp, nhấp từng ngụm và chuyện trò với bạn bè về một năm đã qua? Masala Chai đã chính thức được The Coffee House ra mắt trong giai điệu háo hức của khúc nhạc Giáng sinh an lành, như là một “người bạn dễ thương” mà The Coffee House mang đến để góp mặt cùng bạn trong những câu chuyện hàn huyên, tăng thêm cảm xúc và hương vị giữa những không gian ngập tràn âm thanh và sắc màu để tạm biệt một năm đầy những “nốt lặng”, chào đón năm mới với những điều tươi đẹp hơn. 

Bạn yêu quý, đừng quên rủ bạn bè đến thăm The Coffee House để thưởng thức hương vị ngọt ngào, thơm béo của sữa, cùng vị trà đen mạnh mẽ, vị nồng ấm của những thảo mộc có trong món Trà sữa Masala Chai vào những ngày đặc biệt của mùa Giáng sinh và năm mới này nhé!"
                },
                new News()
                {
                    Id = 6,
                    Thumbnail = "https://file.hstatic.net/1000075078/file/1_1a250d25a8854d07867644995534a34e_grande.jpg",
                    Title = "5 LOẠI TRÀ THẢO MỘC TỐT CHO PHÁI ĐẸP",
                    Content = @"Để có một cơ thể khỏe khoắn và một vóc dáng chuẩn chỉn thì việc ăn uống khoa học và điều độ đóng vai trò rất quan trọng. Trong đó, những sản phẩm từ tự nhiên thường xuyên được nhắc đến và ưu tiên sử dụng. The Coffee House khuyên các chị em nên dùng thường xuyên 05 loại trà thảo mộc quen thuộc dưới đây, vừa thơm ngon dễ uống mà lại chăm sóc cho sức khoẻ và dưỡng nhan cực kỳ hiệu quả. 

Trà hoa cúc 

Hoa cúc là thảo dược có nhiều tác dụng tốt cho sức khỏe. Có vị ngọt, đắng nhẹ, tính mát, nó có tác dụng thanh nhiệt, giải độc, giải cảm, mát gan, làm sáng mắt. Ngoài ra, nếu bạn là người hay bị các chứng mất ngủ, tinh thần căng thẳng, dễ cáu gắt, khó tập trung thì uống trà hoa cúc sẽ giúp làm dịu thần kinh và giúp ngủ ngon hơn.
Đây là loại trà phù hợp cho cả trẻ em và người trưởng thành. Theo các nghiên cứu hiện đại, nó còn có thể giúp kháng khuẩn, kháng siêu vi gây cảm cúm, làm giãn mạch máu, hạ huyết áp, giảm mỡ trong máu. Vậy nên, nếu chị em là những người bận rộn, thường xuyên phải ngồi trước máy tính, ít có thời gian vận động,... nên pha cho mình tách trà hoa cúc để thưởng thức, vừa uống vừa thư giãn nhé!

Trà atiso

Atiso có giá trị dinh dưỡng rất lớn. Chứa khoảng 17 chất khác nhau từ chất xơ, vitamin C, K, B6, Folate, sắt, kẽm,... đồng thời rất giàu chất chống oxy hóa nên sẽ giúp làm chậm quá trình lão hoá, điều mà chị em nào cũng ước ao.
Đặc biệt, trà atiso có khả năng thải độc cho gan, bảo vệ gan khỏi tổn thương và thúc đẩy sự phát triển của tế bào mới. Trong một nghiên cứu lên các đối tượng gan nhiễm mỡ cho thấy việc hấp thụ atiso mỗi ngày trong vòng 2 tháng đã giúp giảm viêm và giảm lượng mỡ hiệu quả. Ngoài ra, dùng trà atiso liên tục trong 6 tuần sẽ làm giảm lượng cholesterol xấu trong cơ thể (khoảng 18.5 ~ 22.9%).

Trà gừng

Vào những ngày đông lạnh, hay những lúc bị cảm và đau họng thì một tách trà gừng sẽ giúp làm ấm cơ thể, chữa cảm lạnh, ngăn ngừa virus hợp bào hô hấp gây bệnh. Ngoài ra, lợi thế lớn nhất của trà gừng là tốt cho dạ dày, điều trị các bệnh về đường tiêu hoá, thậm chí cả ung thư ruột kết.
Một tác dụng bất ngờ của trà gừng nói riêng và các món có chứa gừng đó là chống lại hiệu ứng phá hủy tế bào bởi các gốc tự do, thủ phạm gây nhiều bệnh nan y trong đó có cả bệnh ung thư, tim mạch và đột quỵ. Trong gừng có chứa nhiều hợp chất pararadol và gingerol, có tác dụng tiêu diệt tế bào ung thư, hạn chế khối u di căn sang các tế bào, mô và các bộ phận khỏe mạnh khác.

Trà bạc hà

Trà bạc hà được nhiều người yêu thích bởi vị the mát của nó, kết hợp với vị trái cây thơm ngọt sẽ tạo nên một thức sảng khoái và đầy năng lượng. Hơn thế, trà bạc hà từ lâu được biết đến là loại thức uống thảo dược nóng, tốt cho sức khỏe, nó sẽ giúp bạn cảm thấy thư giãn, cũng như chống lại chứng mất ngủ.
Ngoài ra, trà bạc hà sẽ có nhiều tác dụng khác như cải thiện dòng chảy của mật, cho phép cơ thể tiêu hóa chất béo. Khi sử dụng trà bạc hà còn giúp bạn chữa chứng khó tiêu, các loại rối loạn tiêu hóa khác như buồn nôn, đau bụng, và tiêu chảy. 

Trà xanh

Trà xanh khá quen thuộc và đặc biệt đây là loại trà thảo mộc có rất nhiều công dụng hữu ích đối với các chị em. Trong trà xanh chứa chất chống oxy hóa có hiệu lực mạnh hơn nhiều so với vitamin C và vitamin E, nhờ đó giúp bảo vệ các tế bào không bị phá hủy bởi bệnh ung thư, giảm lượng cholesterol, ngăn ngừa bệnh tim mạch và nguy cơ đột quỵ. 
Đặc biệt, uống trà xanh thường xuyên sẽ giúp phái đẹp làm chậm quá trình lão hóa, mờ các nếp nhăn trên khuôn mặt vì trong trà xanh có chứa polyphenols, một chất chống oxy hóa chống lại các gốc tự do. Ngoài ra, đây là một thức uống tự nhiên có tác dụng hữu hiệu trong việc giảm cân, giúp đốt cháy mỡ thừa và tăng cường khả năng trao đổi chất một cách tự nhiên. Bên cạnh đó, trà xanh sẽ giúp hệ xương khỏe mạnh nhờ chứa florua cao, giảm nguy cơ viêm khớp dạng thấp.

Nếu băn khoăn trong việc lựa chọn sản phẩm trà xanh chất lượng, chị em có thể tham khảo dòng sản phẩm trà xanh Tearoma của The Coffee House. Hương vị trà thanh mát, tinh tế và dễ chịu, có thể lựa chọn mùi vị của đào, sen, nhài nếu chị em yêu thích. Đồng thời, các sản phẩm chế biến từ trà của The Coffee House như trà đào cam sả, trà hạt sen, trà phúc bồn tử, trà macchiato... vừa thơm ngon, lại mang đến các hiệu quả làm đẹp và tốt cho sức khỏe như đã nói trên.
Với những lợi ích của các loại trà thảo mộc đã được The Coffee House chia sẻ, bạn hãy thường xuyên sử dụng nó để đem lại những thay đổi tuyệt vời cho sức khỏe và sắc đẹp của mình nhé. Thế nhưng, cũng đừng quên việc tập luyện thể thao và thiết lập lối sống khoa học. The Coffee House chúc các chị em luôn khoẻ đẹp và rạng ngời!"
                }
                );
            #endregion
            #region Product
            modelBuilder.Entity<Product>().HasData(
                new Product()
                {
                    Id = 1,
                    Name = "Thùng 24 Lon Cà Phê Sữa Đá",
                    Price = 336000,
                    Description = "Cà phê được pha phin truyền thống kết hợp với sữa đặc tạo nên hương vị đậm đà, hài hòa giữa vị ngọt đầu lưỡi và vị đắng thanh thoát nơi hậu vị.",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/24-lon-cpsd_225680_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 1,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 2,
                    Name = "Cà Phê Sữa Đá Hòa Tan",
                    Price = 44000,
                    Description = "Với thiết kế lon cao trẻ trung, hiện đại và tiện lợi, Cà phê sữa đá lon thơm ngon đậm vị của The Coffee House sẽ đồng hành cùng nhịp sống sôi nổi của tuổi trẻ và giúp bạn có được một ngày làm việc đầy hứng khởi.",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/cpsd-3in1_971575_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 1,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 3,
                    Name = "Cà Phê Peak Flavor Hương Thơm Đỉnh Cao (350G)",
                    Price = 90000,
                    Description = "Được rang dưới nhiệt độ vàng, Cà phê Peak Flavor - Hương thơm đỉnh cao lưu giữ trọn vẹn hương thơm tinh tế đặc trưng của cà phê Robusta Đăk Nông và Arabica Cầu Đất. Với sự hòa trộn nhiều cung bậc giữa hương và vị sẽ mang đến cho bạn một ngày mới tràn đầy cảm hứng.",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/peak-plavor-nopromo_715372_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 1,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 4,
                    Name = "Cà Phê Sữa Đá",
                    Price = 90000,
                    Description = "Cà phê được pha phin truyền thống kết hợp với sữa đặc tạo nên hương vị đậm đà, hài hòa giữa vị ngọt đầu lưỡi và vị đắng thanh thoát nơi hậu vị.",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/rich-finish-nopromo_485968.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 1,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 5,
                    Name = "Bạc Sỉu",
                    Price = 32000,
                    Description = "Bạc sỉu chính là 'Ly sữa trắng kèm một chút cà phê'. Thức uống này rất phù hợp những ai vừa muốn trải nghiệm chút vị đắng của cà phê vừa muốn thưởng thức vị ngọt béo ngậy từ sữa.",
                    Photo = "https://file.hstatic.net/1000075078/file/bac-xiu-da_52223c163d7e4bb7921811d9781b43de.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 2,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 6,
                    Name = "Caramel Macchiato Đá",
                    Price = 50000,
                    Description = "Caramel Macchiato sẽ mang đến một sự ngạc nhiên thú vị khi vị thơm béo của bọt sữa, sữa tươi, vị đắng thanh thoát của cà phê Espresso hảo hạng và vị ngọt đậm của sốt caramel được gói gọn trong một tách cà phê.",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/caramel-macchiato_143623_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 2,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 7,
                    Name = "Cà Phê Đá Xay-Lạnh",
                    Price = 59000,
                    Description = "Caramel Macchiato sẽ mang đến một sự ngạc nhiên thú vị khi vị thơm béo của bọt sữa, sữa tươi, vị đắng thanh thoát của cà phê Espresso hảo hạng và vị ngọt đậm của sốt caramel được gói gọn trong một tách cà phê.",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/cf-da-xay-(1)_158038_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 2,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 8,
                    Name = "Trà sữa Oolong Nướng Trân Châu",
                    Price = 55000,
                    Description = @"Trà sữa Oolong nướng Trân Châu – Chuẩn gu đậm đà, uống là mê
Nằm trong bộ sưu tập Trà sữa Nhà đúng gu, Trà sữa Oolong nướng Trân châu là một mảnh ghép được tạo ra để đáp ứng Gu đậm đà của các tín đồ trà sữa.
Trà sữa Oolong nướng mang hương vị chân ái đúng gu đậm đà với trà Oolong được “sao” (nướng) lâu hơn cho hương vị đậm đà, hòa quyện với sữa thơm béo mang đến cảm giác mát lạnh, lưu luyến vị trà sữa đậm đà nơi vòm họng.
The Coffee House tin rằng, khi bạn thưởng thức đúng hương vị trà sữa mà mình thích cũng là cách tận hưởng khoảng khắc tuyệt vời của cuộc sống.
Trà Oolong có tác dụng gì?
Vì Oolong là nguyên liệu chính làm nên thức uống “vạn người mê” này nên khi uống cũng mang đến những lợi ích tuyệt vời cho sức khỏe:
Giúp ngăn ngừa tiểu đường, tốt cho hệ tim mạch, hỗ trợ trong việc giảm cân ,…
Trà Oolong giúp tăng sức đề kháng, giữ cơ thể khỏe mạnh, tinh thần sảng khoái.
Trà Oolong cũng giúp làm dịu dạ dày hiệu quả.
Gọi ngay một ly Trà sữa Oolong nướng thêm Trân châu thôi nào!",
                    Photo = "https://product.hstatic.net/1000075078/product/olong-nuong-tran-chau_017573_400x400_ae76ffea39cd42e4ad8e957678bf4f82_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 3,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 9,
                    Name = "Hồng Trà Sữa Trân Châu",
                    Price = 55000,
                    Description = @"Ngày ngọt ngào hơn khi có Hồng Trà sữa Trân châu
Nằm trong bộ sưu tập Trà sữa Nhà đúng gu, Hồng Trà sữa Trân châu là một mảnh ghép được tạo ra để đáp ứng Gu ngọt ngào của các tín đồ trà sữa.
Hồng Trà sữa Trân châu chuẩn gu ngọt ngào: Hồng trà nguyên lá, sữa thơm ngậy được cân chỉnh với tỉ lệ hoàn hảo, cùng trân châu trắng dai giòn có sẵn để bạn tận hưởng từng ngụm trà sữa ngọt ngào thơm ngậy thiệt đã.
The Coffee House tin rằng, khi bạn thưởng thức đúng hương vị trà sữa mà mình thích cũng là cách thêm chút ngọt ngào cho ngày mới và tận hưởng khoảnh khắc tuyệt vời của cuộc sống.
Lợi ích của việc uống Hồng Trà Sữa
Vì Hồng trà là nguyên liệu chính làm nên thức uống thơm ngon này nên khi uống cũng mang đến những lợi ích tuyệt vời cho sức khỏe:
Giúp tập trung tinh thần, tăng cường trí nhớ. Hợp chất Polyphenol trong hồng còn giúp chắc khỏe xương, ngăn ngừa bệnh loãng xương ở nữ giới.
Theo Đông y Trung Quốc, Hồng trà tốt cho hệ tiêu hóa, lợi tiểu… Ngoài ra, trong Hồng trà còn dồi dào Flavonoid có tác dụng kháng acid, tiêu diệt các gốc tự do giúp ngăn ngừa ung thư và các bệnh tiêm mạch.
Gọi cho mình một Hồng Trà sữa Trân châu The Coffee House để tự thưởng cho bản thân đi nào!",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/tra-nhan-da_484810_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 3,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 10,
                    Name = "Hồng Trà Latte Macchiato",
                    Price = 55000,
                    Description = @"Trà sữa Mắc ca Trân châu trắng – Làn gió mới của các tín đồ trà sữa
Trà sữa Mắc ca Trân châu trắng là sự kết hợp giữa nền trà Oolong hảo hạng thơm lừng, cùng sữa hạt Mắc Ca bổ dưỡng ngọt dịu ít béo, đặc biệt là topping trân châu trắng dai giòn hấp dẫn.
Đây là một 'làn gió mới' mà The Coffee House mang đến cho bạn hương vị mới, trải nghiệm mới cho để bạn đổi vị tươi mới hơn và thú vị hơn. 
The Coffee House tin rằng,
mỗi ngày sẽ là điều tươi mới hơn với sữa hạt mắc ca thơm ngon,
bổ dưỡng quyện cùng nền trà oolong cho vị cân bằng,
ngọt dịu đi kèm cùng Trân châu trắng giòn dai mang lại cảm giác “đã” trong từng ngụm Trà sữa Mắc ca Trân châu trắng.
Đặt một ly và cho The Coffee House biết cảm nhận nhé!",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/hong-tra-latte_618293_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 3,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 11,
                    Name = "Trà Long Nhãn Hạt Chia",
                    Price = 45000,
                    Description = "Caramel Macchiato sẽ mang đến một sự ngạc nhiên thú vị khi vị thơm béo của bọt sữa, sữa tươi, vị đắng thanh thoát của cà phê Espresso hảo hạng và vị ngọt đậm của sốt caramel được gói gọn trong một tách cà phê.",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/tra-nhan-da_484810_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 3,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 12,
                    Name = "Cà Phê Đá Xay-Lạnh",
                    Price = 59000,
                    Description = @"Cà Phê Đá Xay – Phiên bản Upgrade từ ly cà phê sữa quen thuộcNếu là một Fan cà phê chân chính, chắc hẳn bạn sẽ không bỏ qua phiên bản 'upgrade' từ ly cà phê sữa quen thuộc này của The Coffee House.
Thức uống tuy quen mà lạ
Nằm trong dòng thức uống đá xay quen thuộc của Nhà.Dòng thức uống kết hợp sữa tươi,
đá viên với nhiều nguyên liệu khác.Đặc điểm “nhận dạng” dòng thức uống này chính là phần kem tươi xốp mịn bên trên.Cà phê đá xay là món gắn liền với sự xuất hiện của dòng thức uống này ngay từ những ngày đầu xuất hiện tại The Coffee House.Cà phê đá xay với hương vị cà phê sữa quen thuộc,
đi kèm lớp kem và đá xay mang đến cảm giác mát lạnh,
thơm béo thêm phần quyến rũ.
Một ly cà phê đá xay đạt chuẩn Nhà:
Đậm đà hương vị,không bị nhạt do đá quá tan.
Khi uống cảm nhận được hương thơm của cà phê, thơm béo của sữa.
Đá viên được xay nhuyễn mịn, không bị lẫn các viên đá còn to.
Lớp kem tươi xốp, béo và được xịt vừa đủ lên trên ly.
Những lợi ích khi gọi một ly cà phê đá xay
Bên cạnh hương vị thơm ngon khó cưỡng, một ly cà phê đá xay còn là một món quà ngọt ngào cho tâm hồn, bất cứ khi nào bạn cần “kéo mood”. Một ngụm đá xay mát lạnh giúp bạn hoàn toàn tỉnh táo. Lượng cafein trong một ly đá xay không quá nhiều, đặc biệt phù hợp với các bạn yêu thích hương vị cà phê nhưng lại dễ say. 
Cần tỉnh táo, một ly cà phê đá xay nhé!",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/cf-da-xay-(1)_158038_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 4,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 13,
                    Name = "Cookie Đá Xay",
                    Price = 59000,
                    Description = @"Cookie Đá Xay – Món uống phá cách dễ thươngCookie đá xay - món uống phá cách dễ thương đầy mê hoặc từ The Coffee House có gì đặc biệt?
Thức uống “Top trending” của dòng đá xay 
Nằm trong dòng thức uống đá xay (Ice Blended) quen thuộc của Nhà. Dòng thức uống kết hợp sữa tươi, đá viên với nhiều nguyên liệu khác. Đặc điểm “nhận dạng” dòng thức uống này chính là phần kem tươi xốp mịn bên trên. Cookie đá xay với những mẩu bánh cookies giòn rụm kết hợp ăn ý với sữa tươi, kem tươi béo ngọt và đá xay mát lành, đem đến cảm giác lạ miệng gây thích thú và không thể cưỡng lại. Một món uống phá cách dễ thương đầy mê hoặc.
Một ly cookie đá xay đạt chuẩn Nhà:
Đậm đà hương vị, không bị nhạt do đá quá tan.
Khi uống cảm nhận được hương thơm, độ giòn của cookie, thơm béo của sữa.
Đá viên được xay nhuyễn mịn, không bị lẫn các viên đá còn to.
Lớp kem tươi xốp, béo và được xịt vừa đủ lên trên ly.
Từng ngụm thú vị
Không giống các thức uống còn lại, cookie đá xay không chỉ ngon mà còn vui miệng, với từng mẩu cookie giòn thơm kết hợp với đá xay. Với hương vị ngọt nhưng không ngấy, thơm béo vừa phải đây là thức uống thích hợp mỗi khi bạn cần giải nhiệt, cần sự tỉnh táo hay đơn giản một loại thức uống thú vị cho những ngày bỗng nhiên buồn chán.
Order ngay một ly cookie đá xay mát lạnh nhé!",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/Chocolate-ice-blended_183602_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 4,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 14,
                    Name = "Sinh Tố Việt Quất",
                    Price = 59000,
                    Description = @"Sinh tố Việt quất – Uống ngon, uống “đẹp”
Sự phối hợp hợp tinh tế của Barista Nhà
Sinh tố là tên gọi chung của những món trai cây xay. Ở Đây chúng ta có sinh tố việt quất với thành phần chính là mứt việt quất, sữa chua và Foam cheese. Mứt Việt Quất chua thanh, ngòn ngọt, phối hợp nhịp nhàng với dòng sữa chua bổ dưỡng, thêm vào đó là vị béo của Foam cheese được Barista Nhà phối hợp tinh tế. Tất cả tạo món sinh tố thơm ngon mà cả đầu lưỡi và làn da đều thích. 
Thức uống yêu thích của làn da
Với thành phần chính là mứt việt quất - loại trái cây được nhiều yêu thích vì công dụng cho sức khoẻ và đặc biệt làn da. Sinh tố việt quất có nhiều lợi ích có thể kể đến như:
Tốt cho làn da: Vitamin C trong việt quất được xem như là một dưỡng chất cần thiết giúp da tránh khỏi sự lão hóa da dưới tác động của ánh nắng mặt trời, ô nhiễm và khói bụi. 
Những lợi ích đáng kể đối với sức khoẻ: Ngoài vitamin C, trong việt quất còn chứa các loại khoáng chất khác có ích như Vitamin K, Canxi, magie, kali… giúp xương chắc khoẻ, giảm nguy cơ mắc bệnh tiểu đường, tăng cường hệ tiêu hoá và giúp bạn kiểm soát cân nặng nhờ vào chất xơ có trong thành phần.
Tại the Coffee House, mỗi thức uống ra đời không chỉ đảm bảo tiêu chí ngon miệng mà còn quan tâm đến công dụng đối với sức khoẻ. Gọi ngay một ly sinh tố việt quất làm quà cho làn da bạn nhé!",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/sinh-to-viet-quoc_145138_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 4,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 15,
                    Name = "Chocolate Đá Xay",
                    Price = 59000,
                    Description = @"SÔ CÔ LA ĐÁ XAY – Thức uống cho tín đồ hảo ngọt​
Hương vị chocolate quyến rũ
Sữa và kem tươi béo ngọt được “cá tính hoá” bởi vị chocolate đăng đắng. Có thể nói đây là thức uống được sinh ra dành cho các tín đồ hảo ngọt. Giống như các thức uống khác của dòng đá xay như Cà phê đá xay, Matcha đá xay,…một ly Chocolate đá xay đạt chuẩn Barista Nhà phải: 
Đậm đà hương vị, không bị nhạt do đá quá tan.
Khi uống cảm nhận được hương thơm, vị đăng đắng đặc trưng của chocolate, thơm béo của sữa.
Đá viên được xay nhuyễn mịn, không bị lẫn các viên đá còn to.
Lớp kem tươi xốp, béo và được xịt vừa đủ lên trên ly.​
Nguồn năng lượng cho tinh thần
Chocolate là một loại thực phẩm có nhiều công dụng hữu ích, làm hưng phấn tâm trạng do có caffeine, tyrosine và tryptophan. Trong các loại chocolate, chocolate đen được Barista Nhà dùng trong món chocolate đá xay có hàm lượng ca cao cao cùng với chất chống oxy hóa, ít chất béo nên có lợi cho sức khỏe hơn. Chỉ một lượng nhỏ chocolate cũng giúp bạn tràn đầy năng lượng? Không chỉ thế, chocolate còn được các chuyên gia sức khoẻ công nhận do có nhiều lợi ích có thể kể đến như:
Cải thiện sức khỏe tim mạch: Chocolate có thể giúp giảm huyết áo và chống viêm. Theo các nghiên cứu, chocolate có hàm lượng flavonoid cao. Flavonoid có thể kích thích cơ thể tạo ra nhiều oxit nitric hơn, giúp mở rộng và “thư giãn” mạch máu, do đó làm giảm huyết áp.
Giống như trà xanh, chocolate cũng có epicatechin, giúp não chống lại quá trình hình thành các protein kết dích hay các mảng amyloid. Đây là những yếu tố nguy cơ gây ra bệnh Alzheimer.
Có hiểu lầm cho rằng dùng chocolate sẽ gây tăng cân hay các tình trạng sức khỏe khác. Tuy nhiên, theo các nghiên cứu, chỉ cần bạn dùng một lượng nhỏ sẽ không có tác hại xấu đến cơ thể.
Còn có lựa chọn nào tốt hơn cho tinh thần hơn là một ly chocolate đá xay. Order ngay!",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/Chocolate-ice-blended_400940_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 4,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 16,
                    Name = "Mochi Kem Việt Quất",
                    Price = 19000,
                    Description = "Caramel Macchiato sẽ mang đến một sự ngạc nhiên thú vị khi vị thơm béo của bọt sữa, sữa tươi, vị đắng thanh thoát của cà phê Espresso hảo hạng và vị ngọt đậm của sốt caramel được gói gọn trong một tách cà phê.",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/mochi-vietqwuoc_130861_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 5,
                    SupplierId = 2
                },
                new Product()
                {
                    Id = 17,
                    Name = "Mochi Kem Phúc Bồn Tử",
                    Price = 19000,
                    Description = "Caramel Macchiato sẽ mang đến một sự ngạc nhiên thú vị khi vị thơm béo của bọt sữa, sữa tươi, vị đắng thanh thoát của cà phê Espresso hảo hạng và vị ngọt đậm của sốt caramel được gói gọn trong một tách cà phê.",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/mochi-phucbontu_097500_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 5,
                    SupplierId = 2
                },
                new Product()
                {
                    Id = 18,
                    Name = "Mochi Kem Dừa Dứa",
                    Price = 19000,
                    Description = "Caramel Macchiato sẽ mang đến một sự ngạc nhiên thú vị khi vị thơm béo của bọt sữa, sữa tươi, vị đắng thanh thoát của cà phê Espresso hảo hạng và vị ngọt đậm của sốt caramel được gói gọn trong một tách cà phê.",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/mochi-dua_975992_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 5,
                    SupplierId = 2
                },
                new Product()
                {
                    Id = 19,
                    Name = "Mochi Kem Xoài",
                    Price = 19000,
                    Description = "Caramel Macchiato sẽ mang đến một sự ngạc nhiên thú vị khi vị thơm béo của bọt sữa, sữa tươi, vị đắng thanh thoát của cà phê Espresso hảo hạng và vị ngọt đậm của sốt caramel được gói gọn trong một tách cà phê.",
                    Photo = "https://minio.thecoffeehouse.com/image/admin/mochi-xoai_355815_400x400.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 5,
                    SupplierId = 2
                },
                new Product()
                {
                    Id = 20,
                    Name = "Combo 3 Hộp Cà Phê Sữa Đá Hòa Tan Đậm Vị Hộp 18 gói x 16gr",
                    Price = 109000,
                    Description = @"Các chuyên gia The Coffee House tự hào mang đến sản phẩm cà phê 3in1 với vị đậm đà, hương thơm đặc trưng của ly cà phê sữa truyền thống. Hãy khởi đầu ngày mới thật nhiều hứng khởi cùng cà phê 3in1 đậm vị Việt của The Coffee House bạn nhé!
                                    Cafe mang về tiện lợi,phù hợp sử dụng tại nhà,văn phòng hay đi du lịch.Món quà ý nghĩa cho người thân,bạn bè trong các dịp lễ trong năm.",
                    Photo = "https://product.hstatic.net/1000075078/product/combo-3-hop-ca-phe-sua-da-hoa-tan-dam-vi_5bfc3fb23ebd4ec88f17c08794617b96_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 1,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 21,
                    Name = "Combo 3 Hộp Cà Phê Sữa Đá Hòa Tan",
                    Price = 109000,
                    Description = @"Những ưu điểm của cà phê sữa đá hòa tan
Cà phê sữa đá hòa tan được biết đến không những là một thức uống vẫn giữ được hương vị thơm ngon thuần túy như cà phê phin pha với sữa đặc mà sản phẩm này còn mang đến nhiều lợi ích bất ngờ cho các tín đồ yêu ẩm thực.
Tiện lợi
Không ai có thể phủ nhận được sự tiện lợi của cà phê sữa đá hòa tan. Bạn không cần bộ dụng cụ pha chế lỉnh kỉnh, chỉ với những thao tác đơn giản: Đổ cà phê ra ly và chế nước sôi theo tỷ lệ được hướng dẫn trên bao bì, bạn đã có ngay một ly cà phê thơm ngon, chuẩn vị.
Nếu bạn là một tín đồ cà phê chính hiệu thì việc mỗi buổi sáng thức dậy, chỉ cần ngửi mùi hương cà phê cũng đủ mang lại những năng lượng tích cực. Ngoài ra, việc uống cà phê sáng còn có tác dụng ngăn ngừa lão hóa, giảm cân, ngăn ngừa một số bệnh tuổi già,… Với sự hối hả của nhịp sống hiện đại, cà phê sữa đá hòa tan sẽ là lựa chọn của nhiều người.
Cách nhanh nhất thỏa mãn vị giác
Vào những ngày ở nhà bất chợt thèm một ly cà phê hay bạn đang cần một vị cứu tinh để tiếp thêm năng lượng và sự tỉnh táo để làm việc đêm khuya, một gói cà phê sữa đá hòa tan là một sự lựa chọn hoàn hảo để giải tỏa ngay cơn “nghiện” của bạn. Bên cạnh đó, nếu bạn đi phượt đến nên những nơi hẻo lánh thì đây cũng là một người bạn đồng hành tuyệt vời đấy!.
Bắt đầu một ngày mới thật hứng khởi
Nếu bạn là một tín đồ cà phê chính hiệu thì việc mỗi buổi sáng thức dậy, chỉ cần ngửi mùi hương cà phê cũng đủ mang lại những năng lượng tích cực. Ngoài ra, việc uống cà phê sáng còn có tác dụng ngăn ngừa lão hóa, giảm cân, ngăn ngừa một số bệnh tuổi già,… Với sự hối hả của nhịp sống hiện đại, cà phê sữa đá hòa tan sẽ là lựa chọn của nhiều người.
Sáng tạo nhiều món ngon với cà phê sữa đá hòa tan
Là một thức uống hấp dẫn và tiện lợi là thế, cà phê sữa đá còn được các bạn trẻ và bà nội trợ sáng tạo nên rất nhiều món ngon cho gia đình: Sử dụng tạo mùi cho bánh ngọt, râu cau, nguyên liệu chính làm nên cà phê bọt biển siêu hot,...
The Coffee House mời bạn dùng thử cà phê sữa đá hoà tan MỚI, muốn là có, pha ngay tại nhà mà vẫn đúng gu, thơm ngon, đậm vị. Mỗi gói cà phê sữa đá hoà tan là công thức hoàn hảo cho một ly cà phê sữa thơm ngon, chuẩn chỉnh giúp bạn tỉnh táo và hứng khởi cho một ngày làm việc thật hiệu quả.
Còn chần chừ gì mà không đặt thử ngay hôm nay!",
                    Photo = "https://product.hstatic.net/1000075078/product/combo-3cfsd-nopromo_320619_400x400_29c134e6f3a8485d88d5d28a79a82f69_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 1,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 22,
                    Name = "Cà Phê Sữa Đá Hòa Tan Túi 25 x 22gr",
                    Price = 99000,
                    Description = @"Những ưu điểm của cà phê sữa đá hòa tan

Cà phê sữa đá hòa tan được biết đến không những là một thức uống vẫn giữ được hương vị thơm ngon thuần túy như cà phê phin pha với sữa đặc mà sản phẩm này còn mang đến nhiều lợi ích bất ngờ cho các tín đồ yêu ẩm thực.

1. Tiện lợi

Không ai có thể phủ nhận được sự tiện lợi của cà phê sữa đá hòa tan. Bạn không cần bộ dụng cụ pha chế lỉnh kỉnh, chỉ với những thao tác đơn giản: Đổ cà phê ra ly và chế nước sôi theo tỷ lệ được hướng dẫn trên  bao bì, bạn đã có ngay một ly cà phê thơm ngon, chuẩn vị.

Nếu bạn là một tín đồ cà phê chính hiệu thì việc mỗi buổi sáng thức dậy, chỉ cần ngửi mùi hương cà phê cũng đủ mang lại những năng lượng tích cực. Ngoài ra, việc uống cà phê sáng còn có tác dụng ngăn ngừa lão hóa, giảm cân, ngăn ngừa một số bệnh tuổi già,… Với sự hối hả của nhịp sống hiện đại, cà phê sữa đá hòa tan sẽ là lựa chọn của nhiều người.

2. Cách nhanh nhất thỏa mãn vị giác

Vào những ngày ở nhà bất chợt thèm một ly cà phê hay bạn đang cần một vị cứu tinh để tiếp thêm năng lượng và sự tỉnh táo để làm việc đêm khuya, một gói cà phê sữa đá hòa tan là một sự lựa chọn hoàn hảo để giải tỏa ngay cơn “nghiện” của bạn. Bên cạnh đó, nếu bạn đi phượt đến những nơi hẻo lánh thì đây cũng là một người bạn đồng hành tuyệt vời đấy!.

3.. Sáng tạo nhiều món ngon với cà phê sữa đá hòa tan

Là một thức uống hấp dẫn và tiện lợi là thế, cà phê sữa đá còn được các bạn trẻ và bà nội trợ sáng tạo nên rất nhiều món ngon cho gia đình: Sử dụng tạo mùi cho bánh ngọt, râu cau, nguyên liệu chính làm nên cà phê bọt biển siêu hot,...

 The Coffee House mời bạn dùng thử cà phê sữa đá hoà tan MỚI, muốn là có, pha ngay tại nhà mà vẫn đúng gu, thơm ngon, đậm vị. Mỗi gói cà phê sữa đá hoà tan là công thức hoàn hảo cho một ly cà phê sữa thơm ngon, chuẩn chỉnh giúp bạn tỉnh táo và hứng khởi cho một ngày làm việc thật hiệu quả.

Còn chần chừ gì mà không đặt thử ngay hôm nay!

Hướng dẫn sử dụng:
Uống nóng - Hoà 1 gói cà phê sữa đá hoà tan với 100ml nước nóng, khuấy đều và thưởng thức.

Uống đá - Hoà 01 gói cà phê sữa đá hoà tan với 50ml nước nóng và khuấy đều. Cho thêm 100gr đá và thưởng thức.

Bảo quản: Bảo quản sản phẩm nơi khô ráo, thoáng mát, tránh ánh nắng trực tiếp.",
                    Photo = "https://product.hstatic.net/1000075078/product/ca-phe-sua-da-hoa-tan-tui_55b6d640e3594c1b9de982ff6823ebde_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 1,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 23,
                    Name = "Cà Phê Rang Xay Original 1 250gr",
                    Price = 60000,
                    Description = @"Cà phê rang xay Original 1 – Gu đậm truyền thống
Mỗi hạt cà phê ra đời mang trong mình chất nguyên bản,
                    tạo nên khác biệt đầy thú vị.
Trưởng thành từ vùng đất bazan nắng gắt,
                    mưa nhiều,
                    đâm chồi từ đôi bàn tay cần mẫn của người nông dân yêu nghề và rồi cũng từ đôi bàn tay ấy những hạt cà phê đỏ mọng chất lượng tốt nhất được chắt chiu thu hoạch,
                    rang xay tỉ mỉ trước khi về “Nhà”,
                    để gửi đến bạn hương vị đậm đà,
                    tôn vinh chất nguyên bản vốn có của hạt cà phê núi rừng Đắk Lắk.
The Coffee House đã chọn ra những hạt cà phê tốt nhất,
                    kết hợp công nghệ rang xay hiện đại để mang đến bạn gói Cà Phê Original 1 The Coffee House với 100 % Robusta Đăk Lăk giữ trọn vị đậm đà truyền thống thỏa mãn vị giác của tín đồ cà phê Việt.
Bạn đã sẵn sàng để thưởng thức ly cà phê mới từ The Coffee House chưa ? Cùng thử và cho chúng mình biết cảm nhận nhé!",
                    Photo = "https://file.hstatic.net/1000075078/file/ca-phe-rang-xay-original-1-250gr_35d014448d054d63b5e764df85d94f2e.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 1,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 24,
                    Name = "Bạc Sỉu Nóng",
                    Price = 32000,
                    Description = "Bạc sỉu chính là 'Ly sữa trắng kèm một chút cà phê'. Thức uống này rất phù hợp những ai vừa muốn trải nghiệm chút vị đắng của cà phê vừa muốn thưởng thức vị ngọt béo ngậy từ sữa.",
                    Photo = "https://product.hstatic.net/1000075078/product/bac-xiu-nong_fa679e04523c4f138354e852745b223e_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 2,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 25,
                    Name = "Cà phê Sữa Nóng",
                    Price = 35000,
                    Description = "",
                    Photo = "https://product.hstatic.net/1000075078/product/ca-phe-sua-nong_0beecc27870549a5bf016dc6a8fac60e_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 2,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 26,
                    Name = "Cà Phê Đen Nóng",
                    Price = 35000,
                    Description = "",
                    Photo = "https://product.hstatic.net/1000075078/product/ca-phe-sua-nong_0beecc27870549a5bf016dc6a8fac60e_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 2,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 27,
                    Name = "Cà Phê Đen đá",
                    Price = 29000,
                    Description = "",
                    Photo = "https://file.hstatic.net/1000075078/file/ca-phe-den-da_2db07af3c05b4ad5ae04ec1ba04b3351.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 2,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 28,
                    Name = "Cà Phê Đen đá",
                    Price = 29000,
                    Description = "",
                    Photo = "https://file.hstatic.net/1000075078/file/ca-phe-den-da_2db07af3c05b4ad5ae04ec1ba04b3351.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 2,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 29,
                    Name = "Cà Phê Sữa Đá Chai Fresh 250ml",
                    Price = 79000,
                    Description = "",
                    Photo = "https://file.hstatic.net/1000075078/file/ca-phe-sua-da-chai-fresh-250ml_98234cf2b2824b3892168db35402d44a.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 2,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 30,
                    Name = "Trà Đào Cam Sả Chai Fresh 500ml",
                    Price = 109000,
                    Description = "",
                    Photo = "https://file.hstatic.net/1000075078/file/tra-dao-cam-sa-chai-fresh-500ml_34e18021c4b04868aacff6df0388e1f4.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 3,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 31,
                    Name = "Trà Đào Cam Sả Đá",
                    Price = 45000,
                    Description = @"Trà đào cam sả - Thức uống nhất định phải thử 1 lần khi đến The Coffee House
Có thể bạn chưa biết, Trà Đào Cam Sả chính là một trong những món đồ uống được rất nhiều thực khách gọi mỗi khi ghé thăm The Coffee House.
Hương vị tuy quen mà lạ,tuy lạ mà quen.
Khi sáng chế ra thức uống này,Barista tại Nhà đã để lại dấu ấn riêng rất Việt trong món trà đào vốn đã quen thuộc đối với mọi người bằng một nguyên liệu rất đỗi dân dã - Sả,
một loại cây vốn được trồng và sử dụng rộng rãi như là một gia vị thường thấy trong bữa ăn hằng ngày tại các nước châu Á.Sả có hương thơm của chanh,
tính ấm,kết hợp với tính hàn của Cam vàng Mỹ,tạo nên sự kết hợp cân bằng,vừa mang dấu ấn riêng của Việt Nam,vừa có sự pha trộn phương Tây trong ly Trà đào Cam Sả The Coffee House.Vị thanh ngọt của đào Hy Lạp,vị chua dịu của Cam Vàng nguyên vỏ,
vị chát của trà đen tươi được ủ mới mỗi 4 tiếng,cùng hương thơm nồng đặc trưng của sả chính là điểm sáng làm nên sức hấp dẫn của thức uống này.
Bổ sung thêm C,Tăng cường sức đề kháng.Không chỉ hương vị thơm ngon,Trà đào cam sả còn có các tác dụng tuyệt vời cho cơ thể.Có thể kể đến như:
Chống oxy hóa: Trong trà đào có chất chống oxy hóa mạnh mẽ và các hợp chất có thể ngăn ngừa quá trình lão hóa của da.
Tăng cường hệ miễn dịch: Chất chống oxy hóa có trong sả và cam còn góp phần nâng cao hệ thống miễn dịch,
làm giảm căng thẳng và mệt mỏi.Một ly trà đào mỗi ngày sẽ bổ sung lượng lớn vitamin C giúp cơ thể
Giảm cân: Tin vui cho bạn là trà đào có thể hỗ trợ đốt cháy chất béo góp phần và quá trình giảm cân hiệu quả.Vừa là thức uống ngon miệng,
thời thượng lại vừa có thể giảm béo thì quả là tuyệt vời.
Để hiểu vì sao mỗi ngày có đến hơn 20,
000 ly Trà Đào Cam Sả được The Coffee House trao đến tay khách hàng,
order một ly Trà đào cam sả 'trứ danh' để cảm nhận,bạn nhé!",
                    Photo = "https://product.hstatic.net/1000075078/product/tra-dao-cam-xa_668678_400x400_207c526c987c4026877ebae748c62afd_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 3,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 32,
                    Name = "Trà Đào Cam Sả Nóng",
                    Price = 52000,
                    Description = @"Trà đào cam sả - Thức uống nhất định phải thử 1 lần khi đến The Coffee House
Có thể bạn chưa biết, Trà Đào Cam Sả chính là một trong những món đồ uống được rất nhiều thực khách gọi mỗi khi ghé thăm The Coffee House. 
Hương vị tuy quen mà lạ, tuy lạ mà quen
Khi sáng chế ra thức uống này, Barista tại Nhà đã để lại dấu ấn riêng rất Việt trong món trà đào vốn đã quen thuộc đối với mọi người bằng một nguyên liệu rất đỗi dân dã - Sả, một loại cây vốn được trồng và sử dụng rộng rãi như là một gia vị thường thấy trong bữa ăn hằng ngày tại các nước châu Á. Sả có hương thơm của chanh, tính ấm, kết hợp với tính hàn của Cam vàng Mỹ, tạo nên sự kết hợp cân bằng, vừa mang dấu ấn riêng của Việt Nam, vừa có sự pha trộn phương Tây trong ly Trà đào Cam Sả The Coffee House. Vị thanh ngọt của đào Hy Lạp, vị chua dịu của Cam Vàng nguyên vỏ, vị chát của trà đen tươi được ủ mới mỗi 4 tiếng, cùng hương thơm nồng đặc trưng của sả chính là điểm sáng làm nên sức hấp dẫn của thức uống này. 
Bổ sung thêm C, Tăng cường sức đề kháng
Không chỉ hương vị thơm ngon, Trà đào cam sả còn có các tác dụng tuyệt vời cho cơ thể. Có thể kể đến như:
Chống oxy hóa: Trong trà đào có chất chống oxy hóa mạnh mẽ và các hợp chất có thể ngăn ngừa quá trình lão hóa của da. 
Tăng cường hệ miễn dịch: Chất chống oxy hóa có trong sả và cam còn góp phần nâng cao hệ thống miễn dịch, làm giảm căng thẳng và mệt mỏi. Một ly trà đào mỗi ngày sẽ bổ sung lượng lớn vitamin C giúp cơ thể 
Giảm cân: Tin vui cho bạn là trà đào có thể hỗ trợ đốt cháy chất béo góp phần và quá trình giảm cân hiệu quả. Vừa là thức uống ngon miệng, thời thượng lại vừa có thể giảm béo thì quả là tuyệt vời .
Để hiểu vì sao mỗi ngày có đến hơn 20,000 ly Trà Đào Cam Sả được The Coffee House trao đến tay khách hàng, order một ly Trà đào cam sả 'trứ danh' để cảm nhận, bạn nhé!",
                    Photo = "https://product.hstatic.net/1000075078/product/tdcs-nong_288997_400x400_75f2d80e0db2442e9ab752689cbf1001_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 3,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 33,
                    Name = "Trà Hạt Sen Đá",
                    Price = 45000,
                    Description = @"Trà hạt sen – Sự tinh tế của những điều đơn giản
Nền trà oolong hảo hạng kết hợp cùng hạt sen tươi, bùi bùi và lớp foam cheese béo ngậy. Thức uống gây thương nhớ tại The Coffee House có gì đặc biệt?
Sự kết hợp tinh tế của những nguyên liệu giản đơn 
Sự kết hợp của Trà Oolong thanh mát với hương thơm nhẹ nhàng cùng vị nồng hậu của Hạt sen tươi mềm có vị ngọt, bùi. Đặc biệt, lớp kem sữa phô mai được phủ lên bề mặt ly sẽ cho bạn một trải nghiệm cân bằng hơn về hương vị.
Hình thức cực kỳ bắt mắt
Một ly trà hạt sen chuẩn Nhà, bạn có thể nhìn thấy rõ ràng 3 lớp riêng biệt. Lớp đầu tiên là lớp “milk foam”béo ngậy, lớp ở giữa là trà oolong màu vàng sáng và lắng đọng ở lớp đáy là những “hạt ngọc” sen vàng.
Trà hạt sen bồi bổ sức khoẻ:
Trà Oolong – thảo dược quý giúp bạn bồi bổ sức khoẻ: Trong trà ô long có chứa rất nhiều vitamin, khoáng chất và những chất chống oxy hóa rất tốt. 
 Hạt sen  cung cấp một lượng lớn gluxit, lipit, canxi, chất xơ và các vitamin mang đến nhiều lợi ích ch người dùng từ tác dụng chống viêm, giúp an thần, đến công dụng đẹp da, chống lão hoá, điều hoà cholesterol và đường huyết. 
Một loại thức uống phù hợp cho cả buổi sớm ban mai hay chiều tối với nhiều công dụng tuyệt vời chơ sức khoẻ, bạn đã thử chưa?",
                    Photo = "https://product.hstatic.net/1000075078/product/tra-sen_905594_400x400_c82fd1b3786d45f380cf4c15e9af7ab9_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 3,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 34,
                    Name = "Trà Hạt Sen Nóng",
                    Price = 52000,
                    Description = @"Trà hạt sen – Sự tinh tế của những điều đơn giản
Nền trà oolong hảo hạng kết hợp cùng hạt sen tươi, bùi bùi và lớp foam cheese béo ngậy. Thức uống gây thương nhớ tại The Coffee House có gì đặc biệt?
Sự kết hợp tinh tế của những nguyên liệu giản đơn 
Sự kết hợp của Trà Oolong thanh mát với hương thơm nhẹ nhàng cùng vị nồng hậu của Hạt sen tươi mềm có vị ngọt, bùi. Đặc biệt, lớp kem sữa phô mai được phủ lên bề mặt ly sẽ cho bạn một trải nghiệm cân bằng hơn về hương vị.
Hình thức cực kỳ bắt mắt
Một ly trà hạt sen chuẩn Nhà, bạn có thể nhìn thấy rõ ràng 3 lớp riêng biệt. Lớp đầu tiên là lớp “milk foam”béo ngậy, lớp ở giữa là trà oolong màu vàng sáng và lắng đọng ở lớp đáy là những “hạt ngọc” sen vàng.
Trà hạt sen bồi bổ sức khoẻ:
Trà Oolong – thảo dược quý giúp bạn bồi bổ sức khoẻ: Trong trà ô long có chứa rất nhiều vitamin, khoáng chất và những chất chống oxy hóa rất tốt. 
 Hạt sen  cung cấp một lượng lớn gluxit, lipit, canxi, chất xơ và các vitamin mang đến nhiều lợi ích ch người dùng từ tác dụng chống viêm, giúp an thần, đến công dụng đẹp da, chống lão hoá, điều hoà cholesterol và đường huyết. 
Một loại thức uống phù hợp cho cả buổi sớm ban mai hay chiều tối với nhiều công dụng tuyệt vời chơ sức khoẻ, bạn đã thử chưa?",
                    Photo = "https://product.hstatic.net/1000075078/product/tra-sen-nong_025153_400x400_b7b0ec5844c546d2a6d139a65be8aa0a_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 3,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 35,
                    Name = "Hồng Trà Sữa Nóng",
                    Price = 50000,
                    Description = @"Ngày ngọt ngào hơn khi có Hồng Trà sữa Trân châu
Nằm trong bộ sưu tập Trà sữa Nhà đúng gu, Hồng Trà sữa Trân châu là một mảnh ghép được tạo ra để đáp ứng Gu ngọt ngào của các tín đồ trà sữa.
Hồng Trà sữa Trân châu chuẩn gu ngọt ngào: Hồng trà nguyên lá, sữa thơm ngậy được cân chỉnh với tỉ lệ hoàn hảo, cùng trân châu trắng dai giòn có sẵn để bạn tận hưởng từng ngụm trà sữa ngọt ngào thơm ngậy thiệt đã.
The Coffee House tin rằng, khi bạn thưởng thức đúng hương vị trà sữa mà mình thích cũng là cách thêm chút ngọt ngào cho ngày mới và tận hưởng khoảnh khắc tuyệt vời của cuộc sống.
Lợi ích của việc uống Hồng Trà Sữa
Vì Hồng trà là nguyên liệu chính làm nên thức uống thơm ngon này nên khi uống cũng mang đến những lợi ích tuyệt vời cho sức khỏe:
Giúp tập trung tinh thần, tăng cường trí nhớ. Hợp chất Polyphenol trong hồng còn giúp chắc khỏe xương, ngăn ngừa bệnh loãng xương ở nữ giới.
Theo Đông y Trung Quốc, Hồng trà tốt cho hệ tiêu hóa, lợi tiểu… Ngoài ra, trong Hồng trà còn dồi dào Flavonoid có tác dụng kháng acid, tiêu diệt các gốc tự do giúp ngăn ngừa ung thư và các bệnh tiêm mạch.
Gọi cho mình một Hồng Trà sữa Trân châu The Coffee House để tự thưởng cho bản thân đi nào!",
                    Photo = "https://product.hstatic.net/1000075078/product/hong-tra-sua-nong_941687_400x400_dfe4a7a23ecd4ca49822fc553909828a_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 3,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 36,
                    Name = "Hồng Trà Sữa Trân Châu",
                    Price = 55000,
                    Description = @"Ngày ngọt ngào hơn khi có Hồng Trà sữa Trân châu
Nằm trong bộ sưu tập Trà sữa Nhà đúng gu, Hồng Trà sữa Trân châu là một mảnh ghép được tạo ra để đáp ứng Gu ngọt ngào của các tín đồ trà sữa.
Hồng Trà sữa Trân châu chuẩn gu ngọt ngào: Hồng trà nguyên lá, sữa thơm ngậy được cân chỉnh với tỉ lệ hoàn hảo, cùng trân châu trắng dai giòn có sẵn để bạn tận hưởng từng ngụm trà sữa ngọt ngào thơm ngậy thiệt đã.
The Coffee House tin rằng, khi bạn thưởng thức đúng hương vị trà sữa mà mình thích cũng là cách thêm chút ngọt ngào cho ngày mới và tận hưởng khoảnh khắc tuyệt vời của cuộc sống.
Lợi ích của việc uống Hồng Trà Sữa
Vì Hồng trà là nguyên liệu chính làm nên thức uống thơm ngon này nên khi uống cũng mang đến những lợi ích tuyệt vời cho sức khỏe:
Giúp tập trung tinh thần, tăng cường trí nhớ. Hợp chất Polyphenol trong hồng còn giúp chắc khỏe xương, ngăn ngừa bệnh loãng xương ở nữ giới.
Theo Đông y Trung Quốc, Hồng trà tốt cho hệ tiêu hóa, lợi tiểu… Ngoài ra, trong Hồng trà còn dồi dào Flavonoid có tác dụng kháng acid, tiêu diệt các gốc tự do giúp ngăn ngừa ung thư và các bệnh tiêm mạch.
Gọi cho mình một Hồng Trà sữa Trân châu The Coffee House để tự thưởng cho bản thân đi nào!",
                    Photo = "https://product.hstatic.net/1000075078/product/hong-tra-sua-tran-chau_326977_400x400_30c57e67d1104ae4b7313cf5d40f3e87_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 3,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 37,
                    Name = "Trà Đen Macchiato",
                    Price = 50000,
                    Description = @"Trà đen Macchiato – Chuẩn vị thơm ngon
Trà đen được ủ mới mỗi ngày, giữ nguyên được vị chát mạnh mẽ đặc trưng của lá trà, phủ bên trên là lớp Macchiato 'homemade' bồng bềnh quyến rũ vị phô mai mặn mặn mà béo béo.
Trà đen Macchiato là một trong những thức uống được các bạn trẻ rất yêu thích của The Coffee House.Nếu bạn nào đã thử qua,
chắc chắn sẽ phải siêu lòng ngay bởi sự hòa quyện giữa vị đắng vừa phải của trà,
nhấp từng ngụm bạn sẽ cảm nhận vị ngon khó cưỡng của lớp kem béo ngậy của Macchiato.
Uống Trà Đen Macchiato như thế nào mới đúng cách ?
Khi uống Trà đen Macchiato, bạn thường uống theo cách nào nhất ?
Giữ nguyên nắp và uống.
Tháo nắp ra để lớp sữa chạm vào môi và ngửi được hương thơm.
Lắc lên trước để trà và sữa quyện vào nhau.
Tuy nhiên,cách uống ngon nhất vẫn là cách bạn chọn! Bạn thường thưởng thức trà đen theo cách nào, chia sẻ với The Coffee House với nhé!",
                    Photo = "https://product.hstatic.net/1000075078/product/tra-den-matchiato_430281_400x400_88f47618a6b84c2d9ad633d3298fb390_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 3,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 38,
                    Name = "Trà Sữa Mắc Ca Trân Châu Trắng",
                    Price = 50000,
                    Description = @"Trà sữa Mắc ca Trân châu trắng – Làn gió mới của các tín đồ trà sữa
Trà sữa Mắc ca Trân châu trắng là sự kết hợp giữa nền trà Oolong hảo hạng thơm lừng, cùng sữa hạt Mắc Ca bổ dưỡng ngọt dịu ít béo, đặc biệt là topping trân châu trắng dai giòn hấp dẫn.
Đây là một 'làn gió mới' mà The Coffee House mang đến cho bạn hương vị mới, trải nghiệm mới cho để bạn đổi vị tươi mới hơn và thú vị hơn. 
The Coffee House tin rằng,
mỗi ngày sẽ là điều tươi mới hơn với sữa hạt mắc ca thơm ngon,
bổ dưỡng quyện cùng nền trà oolong cho vị cân bằng,
ngọt dịu đi kèm cùng Trân châu trắng giòn dai mang lại cảm giác “đã” trong từng ngụm Trà sữa Mắc ca Trân châu trắng.
Đặt một ly và cho The Coffee House biết cảm nhận nhé!",
                    Photo = "https://product.hstatic.net/1000075078/product/tra-sua-mac-ca_377522_400x400_123ab6b1593d4e5c83776a54f6862bbd_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 3,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 39,
                    Name = "Đào Việt Quất Đá Xay",
                    Price = 59000,
                    Description = @"Đào Việt Quất Đá Xay –  Sự Pha Trộn Hoàn Hảo Giữa Đào & Việt Quất
Một trong những món đá xay phá cách được yêu thích nhất tại The Coffee House
Sự pha trộn tinh tế
Là phiên bản “upgrade” từ ly trà đào quen thuộc của nhà, tuy là “em út” trong giá đình đá xay nhưng món đào việt quất đá xay lại nhanh chóng quen mặt, sau thời gian thử thách đã chính thức từ thức uống theo mùa trở thành món uống trên menu chính thức của Nhà. Được khoác lên mình ngoại hình xịn sò, mang đến sự phấn khích và tươi mát ngay từ cái nhìn đầu tiên, ở giữa là vịđào huyền thoại ngọt thanh, lớp việt quất chua ngọt vui miệng ở tầng cuối, đặc biệt kéo theo lớp kem bồng bềnh béo ngậy trên cùng nữa mang đến hương vị kích thích vị giác đầy lôi cuốn và khoan khoái ngay từ ngụm đầu tiên.
Ngon miệng, ngon mắt lại còn chứa nhiều vitamin
Quả đào chứa nguồn vitamin dồi dào, đặc biệt là những loại vitamin thiết yếu. Các loại vitamin và khoáng chất có thể kể đến như: Vitamin K, canxi, phốt pho, đồng, kẽm. Các vitamin & khoáng chất này có nhiều công dụng tốt cho sức khỏe, đặc biệt là khả năng cải thiện thị lực, các vấn đề về xương và thần kinh.
Mứt việt quất chứa vitamin C, Vitamin K, Canxi, magie, kali… giúp xương chắc khoẻ, giảm nguy cơ mắc bệnh tiểu đường, tăng cường hệ tiêu hoá và giúp bạn kiểm soát cân nặng nhờ vào chất xơ có trong thành phần. 
Đặc biệt cả 2 loại quả đều có công dụng giúp da tránh khỏi sự lão hoá. 
Nắng nóng, sức lực làm việc bỗng nhiên bị “bốc hơi”, tự nhiên thèm một thứ gì đó hạ hỏa và lên tinh thần ngay lúc này, order một ly đào việt quất đá xay và thưởng thức nhé!",
                    Photo = "https://product.hstatic.net/1000075078/product/daovietquat_033985_400x400_20bfc56b971c47dca8734514a4765675_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 4,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 40,
                    Name = "Chanh Sả Đá Xay",
                    Price = 49000,
                    Description = @"CHANH SẢ ĐÁ XAY – Thức uống thanh mát, giải nhiệt
Sự kết hợp của các nguyên liệu mộc mạc
Nước cốt sả (sả ngâm) xay cùng với chanh tươi, thêm vào chút đường tạo nên thức ống với hương vị thanh chua, ngọt ngọt, thơm nhẹ thoảng nhẹ mùi sả. Tuy thành phần đơn giản nhưng để có một ly chanh sả đá xay đúng điệu cùng cần tay nghề khéo léo của Barista Nhà từ sự kết hợp giữa các thành phần đến độ nhuyễn mịn của đá xay. 
Thức uống giải nhiệt cơ thể
Tuy là những nguyên liệu đơn giản nhưng sả và chanh lại có những công dụng cực kỳ tốt đối với sức khoẻ:
Sả rất tốt cho hệ tiêu hoá, hỗ trợ tốt cho hệ thần kinh, có tính kháng viêm và đăc biệt tốt cho cho người muốn giảm cân do sả có tác dụng đốt cháy mỡ thừa, làm thúc đẩy quá trình trao đổi giúp máu lưu thông tốt hơn.
Trong khi đó Chanh lại là loại quả cung cấp rất nhiều vitamin C, chất xơ và các hợp chất thực vật có lợi khác nhau. Những chất dinh dưỡng này đem lại một số lợi ích cho sức khỏe. Trên thực tế, chanh có thể hỗ trợ sức khỏe tim mạch, kiểm soát cân nặng và sức khỏe hệ tiêu hóa.
Gọi một ly chanh sả đá xay để giải nhiệt ngay thôi nào!",
                    Photo = "https://product.hstatic.net/1000075078/product/chanh-sa-da-xay_170980_400x400_2c8af8d0cfff43b78a93924d28a1215f_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 4,
                    SupplierId = 1
                },
                new Product()
                {
                    Id = 41,
                    Name = "Mousse Gấu Chocolate",
                    Price = 39000,
                    Description = "Với vẻ ngoài đáng yêu và hương vị ngọt ngào, thơm béo nhất định bạn phải thử ít nhất 1 lần.",
                    Photo = "https://product.hstatic.net/1000075078/product/mousse-gau-chocolate_e25675d96c504dacb88f4e56a54ab48b_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 5,
                    SupplierId = 2
                },
                new Product()
                {
                    Id = 42,
                    Name = "Mousse Passion Cheese",
                    Price = 29000,
                    Description = "Tận hưởng chiếc bánh mát lạnh với sự kết hợp hoàn hảo của vị béo ngậy của nhân kem phô mai, cân bằng với vị chua thanh từ chanh dây.",
                    Photo = "https://product.hstatic.net/1000075078/product/5d92fbc79e12b47db8cabbd0_chanh-day_994413_400x400_333c9a63c0f74d41a11333c637127328_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 5,
                    SupplierId = 2
                },
                new Product()
                {
                    Id = 43,
                    Name = "Mousse Red Velvet",
                    Price = 29000,
                    Description = "Bánh nhiều lớp được phủ lớp kem bên trên bằng Cream cheese.",
                    Photo = "https://product.hstatic.net/1000075078/product/mousse-red-velvet_e732a3a0ffa44f589c8455ab75bc0f60_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 5,
                    SupplierId = 2
                },
                new Product()
                {
                    Id = 44,
                    Name = "Mousse Tiramisu",
                    Price = 32000,
                    Description = "Hương vị dễ ghiền được tạo nên bởi chút đắng nhẹ của cà phê, lớp kem trứng béo ngọt dịu hấp dẫn.",
                    Photo = "https://file.hstatic.net/1000075078/file/mousse-tiramisu_5ee1d908f0574cfabb917f72aac288ad.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 5,
                    SupplierId = 2
                },
                new Product()
                {
                    Id = 45,
                    Name = "Bánh Mì Que Pate",
                    Price = 12000,
                    Description = "Vỏ bánh mì giòn tan, kết hợp với lớp nhân pate béo béo đậm đà sẽ là lựa chọn lý tưởng nhẹ nhàng để lấp đầy chiếc bụng đói , cho 1 bữa sáng - trưa - chiều - tối của bạn thêm phần thú vị.",
                    Photo = "https://product.hstatic.net/1000075078/product/banhmique_056830_400x400_8b3112e96ce949be82b70c4b4ab994c2_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 5,
                    SupplierId = 2
                },
                new Product()
                {
                    Id = 46,
                    Name = "Bánh Mì VN Thịt Nguội",
                    Price = 29000,
                    Description = "Gói gọn trong ổ bánh mì Việt Nam, là từng lớp chả, từng lớp jambon hòa quyện cùng bơ và pate thơm lừng, thêm dưa rau cho bữa sáng đầy năng lượng.",
                    Photo = "https://product.hstatic.net/1000075078/product/banh-mi-vn-thit-nguoi_5e777a62ef2841bd9a5baf85c6970166_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 5,
                    SupplierId = 2
                },
                new Product()
                {
                    Id = 47,
                    Name = "Chà Bông Phô Mai",
                    Price = 32000,
                    Description = "Chiếc bánh với lớp phô mai vàng sánh mịn bên trong, được bọc ngoài lớp vỏ xốp mềm thơm lừng. Thêm lớp chà bông mằn mặn hấp dẫn bên trên.",
                    Photo = "https://product.hstatic.net/1000075078/product/cha-bong-pho-mai_204282_400x400_8ff5f28305f547c287685bb01a8836fd_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 5,
                    SupplierId = 2
                },
                new Product()
                {
                    Id = 48,
                    Name = "Croissant Trứng Muối",
                    Price = 35000,
                    Description = "Croissant trứng muối thơm lừng, bên ngoài vỏ bánh giòn hấp dẫn bên trong trứng muối vị ngon khó cưỡng.",
                    Photo = "https://file.hstatic.net/1000075078/file/croissant-trung-muoi_1bca7334056c41afbf3675dccf8b744e.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 5,
                    SupplierId = 2
                },
                new Product()
                {
                    Id = 49,
                    Name = "Gà Xé Lá Chanh",
                    Price = 29000,
                    Description = "Thịt gà được xé tơi, mang hương vị mặn, ngọt, cay cay quyện nhau vừa chuẩn, thêm chút thơm thơm thơm từ lá chanh sấy khô giòn giòn xua tan ngay cơn buồn miệng.",
                    Photo = "https://product.hstatic.net/1000075078/product/kho-ga-la-chanh_995862_400x400_144de2ed4ca449d8b82bc558155e4b2f_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 5,
                    SupplierId = 2
                },
                new Product()
                {
                    Id = 50,
                    Name = "Mít Sấy",
                    Price = 20000,
                    Description = "Mít sấy khô vàng ươm, giòn rụm, giữ nguyên được vị ngọt lịm của mít tươi.",
                    Photo = "https://product.hstatic.net/1000075078/product/mit-say_666228_400x400_062954ca87384c13ab5044766afbec56_master.jpg",
                    CreatedDate = DateTime.Now,
                    Size = 0,
                    ProductTypeId = 5,
                    SupplierId = 2
                }
                );
            #endregion
            #region ProductType
            modelBuilder.Entity<ProductType>().HasData(
                new ProductType()
                {
                    Id = 1,
                    Name = "Cà phê gói - Uống liền",
                    Description = "Cà phê đóng gói",
                    Photo = "https://minio.thecoffeehouse.com/image/tch-web-order/category-thumbnails/ca-phe-goi-ca-phe-uong-lien.png",
                },
                new ProductType()
                {
                    Id = 2,
                    Name = "Cà phê pha",
                    Description = "Cà phê pha",
                    Photo = "https://minio.thecoffeehouse.com/image/tch-web-order/category-thumbnails/cg_coffee_web.png",
                },
                new ProductType()
                {
                    Id = 3,
                    Name = "Trà Trái Cây - Trà sữa",
                    Description = "Trà Trái Cây - Trà sữa",
                    Photo = "https://minio.thecoffeehouse.com/image/tch-web-order/category-thumbnails/tra-trai-cay-tra-sua.png",
                },
                new ProductType()
                {
                    Id = 4,
                    Name = "Đá xay - Choco - Matcha",
                    Description = "Đá xay",
                    Photo = "https://minio.thecoffeehouse.com/image/tch-web-order/category-thumbnails/da-xa.png",
                },
                new ProductType()
                {
                    Id = 5,
                    Name = "Bánh - snack",
                    Description = "Bánh",
                    Photo = "https://minio.thecoffeehouse.com/image/tch-web-order/category-thumbnails/banh-snack.png",
                },
                new ProductType()
                {
                    Id = 6,
                    Name = "Bộ sưu tập - quà tặng",
                    Description = "Đồ lưu niệm",
                    Photo = "https://minio.thecoffeehouse.com/image/tch-web-order/category-thumbnails/bo-suu-tap-qua-tang.png",
                }
                );
            #endregion
            #region Role
            modelBuilder.Entity<Role>().HasData(
                new Role()
                {
                    Id = 1,
                    RoleName = "Admin",
                    Description = "Someone whose job is to control the operation of all stores."
                },
                new Role()
                {
                    Id = 2,
                    RoleName = "Manager",
                    Description = "The person who is responsible for managing specific store."
                },
                new Role()
                {
                    Id = 3,
                    RoleName = "Staff",
                    Description = "The people who work for a store."
                },
                new Role()
                {
                    Id = 4,
                    RoleName = "Shipper",
                    Description = "A person whose job is to organize the sending of goods to customer."
                }
                );
            #endregion
            #region ShoppingCart
            modelBuilder.Entity<ShoppingCart>().HasData(
                new ShoppingCart()
                {
                    Id = 1,
                    CustomerId = 1,
                },
                new ShoppingCart()
                {
                    Id = 2,
                    CustomerId = 2,
                },
                new ShoppingCart()
                {
                    Id = 3,
                    CustomerId = 3,
                },
                new ShoppingCart()
                {
                    Id = 4,
                    CustomerId = 4,
                }
                );
            #endregion
            #region ShoppingCart_Product
            modelBuilder.Entity<ShoppingCart_Product>().HasData(
                new ShoppingCart_Product()
                {
                    ShoppingCartId = 1,
                    ProductId = 1,
                },
                new ShoppingCart_Product()
                {
                    ShoppingCartId = 1,
                    ProductId = 2,
                },
                new ShoppingCart_Product()
                {
                    ShoppingCartId = 2,
                    ProductId = 1,
                },
                new ShoppingCart_Product()
                {
                    ShoppingCartId = 2,
                    ProductId = 3,
                },
                new ShoppingCart_Product()
                {
                    ShoppingCartId = 3,
                    ProductId = 3,
                },
                new ShoppingCart_Product()
                {
                    ShoppingCartId = 3,
                    ProductId = 2,
                },
                new ShoppingCart_Product()
                {
                    ShoppingCartId = 4,
                    ProductId = 1,
                },
                new ShoppingCart_Product()
                {
                    ShoppingCartId = 4,
                    ProductId = 4,
                }
                );
            #endregion    
            #region Store
            modelBuilder.Entity<Store>().HasData(
                new Store
                {
                    Id = 1,
                    StoreName = "HCM Trần Quang Khải",
                    Address = "180 Trần Quang Khải, Quận 1, Hồ Chí Minh",
                    Country = "Việt Nam",
                    Phone = "00000000001",
                    Description = "Quán café và sách tọa lạc tại Quận 1 ở Thành phồ Hồ Chí Minh, Việt Nam",
                    District = "1",
                    Photo = "https://lh5.googleusercontent.com/p/AF1QipNrmu9PP94S52FOmcYEJgRfx-ryY-V96JNSVLhL=w408-h291-k-no",
                    LinkGG = "https://goo.gl/maps/ftFGRq9FNoJy6p887"
                },
                new Store
                {
                    Id = 2,
                    StoreName = "HCM Nguyễn Thái Bình",
                    Address = "141 Nguyễn Thái Bình, Quận 1, Hồ Chí Minh",
                    Country = "Việt Nam",
                    Phone = "00000000002",
                    Description = "Quán café và sách tọa lạc tại Quận 1 ở Thành phồ Hồ Chí Minh, Việt Nam",
                    District = "1",
                    Photo = "https://lh5.googleusercontent.com/p/AF1QipPtXSa8-F-grDLfbTWBytLwupn9cuJ5oCTuWd3F=w408-h291-k-no",
                    LinkGG = "https://goo.gl/maps/WWzyyXyiUDGP875c8"
                },
                new Store
                {
                    Id = 3,
                    StoreName = "HCM The Hub - Điện Biên Phủ",
                    Address = "195/10 Điện Biên Phủ, Phường 15, Quận Bình Thạnh, Tp.HCM",
                    Country = "Việt Nam",
                    Phone = "00000000003",
                    Description = "Quán café và sách tọa lạc tại Quận Bình Thạnh ở Thành phồ Hồ Chí Minh, Việt Nam",
                    District = "Bình Thạnh",
                    Photo = "https://file.hstatic.net/1000075078/file/_dsc7394_756ced0f2a8d4e189d4b3b7f3b15dc68.jpeg",
                    LinkGG = "https://goo.gl/maps/anBoZoESsXUPefXT7"
                },
                new Store
                {
                    Id = 4,
                    StoreName = "HCM Nguyễn Gia Trí (Đường D2)",
                    Address = "157 Nguyễn Gia Trí, Phường 25, Bình Thạnh, Thành phố Hồ Chí Minh, Vietnam",
                    Country = "Việt Nam",
                    Phone = "00000000004",
                    Description = "Quán café và sách tọa lạc tại Quận Bình Thạnh ở Thành phồ Hồ Chí Minh, Việt Nam",
                    District = "Bình Thạnh",
                    Photo = "https://file.hstatic.net/1000075078/file/hcm-nguyen-gia-tri3_88104990bcf24980960e3c8fdeceec8b.jpg",
                    LinkGG = "https://goo.gl/maps/XZ8hhmUKDJYw3BqU6"
                },
                new Store
                {
                    Id = 5,
                    StoreName = "HCM Nguyễn Xí",
                    Address = "184 Nguyễn Xí, Phường 26, Bình Thạnh, Thành phố Hồ Chí Minh",
                    Country = "Việt Nam",
                    Phone = "00000000005",
                    Description = "Quán café và sách tọa lạc tại Quận Bình Thạnh ở Thành phồ Hồ Chí Minh, Việt Nam",
                    District = "Bình Thạnh",
                    Photo = "https://file.hstatic.net/1000075078/file/hcm-nguyen-xi1_b8bb3a5f3381470e90405508724a7a16.jpg",
                    LinkGG = "https://goo.gl/maps/s7ZiN5WVHFEU11c46"
                },
                new Store
                {
                    Id = 6,
                    StoreName = "HCM Phan Văn Trị",
                    Address = "190 Phan Văn Trị, Phường 11, Bình Thạnh, Thành phố Hồ Chí Minh, Vietnam",
                    Country = "Việt Nam",
                    Phone = "00000000006",
                    Description = "Quán café và sách tọa lạc tại Quận Bình Thạnh ở Thành phồ Hồ Chí Minh, Việt Nam",
                    District = "Bình Thạnh",
                    Photo = "https://file.hstatic.net/1000075078/file/hcm-phan-van-tri-31_dea2c72cc4e24431abceee9b11a67ba2.jpg",
                    LinkGG = "https://goo.gl/maps/xcrHU7rT2hxaB3LR9"
                },
                new Store
                {
                    Id = 7,
                    StoreName = "HCM Cao Thắng",
                    Address = "175B Cao Thắng, Phường 12, Quận 10, Hồ Chí Minh",
                    Country = "Việt Nam",
                    Phone = "00000000007",
                    Description = "Quán café và sách tọa lạc tại Quận 10 ở Thành phố Hồ Chí Minh, Việt Nam",
                    District = "10",
                    Photo = "https://file.hstatic.net/1000075078/file/hcm-cao-thang-21_8fba3187121f4a1da8306a144c1d509e.jpg",
                    LinkGG = "https://goo.gl/maps/WkzbFYqBoLQ96LPm9"
                },
                new Store
                {
                    Id = 8,
                    StoreName = "HCM Sư Vạn Hạnh",
                    Address = "798 Sư Vạn Hạnh, Phường 12, Quận 10, Hồ Chí Minh",
                    Country = "Việt Nam",
                    Phone = "00000000008",
                    Description = "Quán café và sách tọa lạc tại Quận 10 ở Thành phố Hồ Chí Minh, Việt Nam",
                    District = "10",
                    Photo = "https://file.hstatic.net/1000075078/file/hcm-su-van-hanh1_dc98609ed5264bbd941fff58c28e2ea6.jpg",
                    LinkGG = "https://goo.gl/maps/dvMsEA3Q1FKyU7Qr6"
                },
                new Store
                {
                    Id = 9,
                    StoreName = "HCM Nguyễn Duy Trinh",
                    Address = "670 Nguyễn Duy Trinh, Bình Trưng Đông, Quận 2, Thành phố Thủ Đức",
                    Country = "Việt Nam",
                    Phone = "00000000009",
                    Description = "Quán café và sách tọa lạc tại Quận 2 ở Thành phố Thủ Đức, Việt Nam",
                    District = "2",
                    Photo = "https://file.hstatic.net/1000075078/file/hcm-nguyen-duy-trinh1_dd549b9380cd461989e27b53dd7954c1.jpg",
                    LinkGG = "https://goo.gl/maps/vVSMkvj1AneVZeWL7"
                },
                new Store
                {
                    Id = 10,
                    StoreName = "HCM Hoàng Diệu",
                    Address = "66E Hoàng Diệu 2, P.Linh Trung, Thủ Đức, Hồ Chí Minh",
                    Country = "Việt Nam",
                    Phone = "00000000010",
                    Description = "Quán café và sách tọa lạc tại Quận Thư Đức ở Thành phố Thủ Đức, Việt Nam",
                    District = "2",
                    Photo = "https://file.hstatic.net/1000075078/file/hcm-hoang-dieu-21_d400dc36e41e4bbabbb442bbe15f9d99.jpg",
                    LinkGG = "https://goo.gl/maps/HSi4jDTYMuQ1dZYg6"
                }
                );
            #endregion
            #region Supplier
            modelBuilder.Entity<Supplier>().HasData(
                new Supplier()
                {
                    Id = 1,
                    Name = "Nhà phân phối sản phẩm đồ uống",
                    Phone = "01122334455",
                    Address = "Quận 1",
                    City = "Hồ Chí Minh",
                    Country = "Việt Nam",
                    Description = "Chuyên phân phối nguyên liệu, sản phầm đồ uống",
                    Url = "supplier1.com"
                },
                new Supplier()
                {
                    Id = 2,
                    Name = "Nhà phân phối sản phẩm đồ ăn",
                    Phone = "11122334455",
                    Address = "Quận 1",
                    City = "Hồ Chí Minh",
                    Country = "Việt Nam",
                    Description = "Chuyên phân phối nguyên liệu, sản phầm đồ ăn",
                    Url = "supplier2.com"
                },
                new Supplier()
                {
                    Id = 3,
                    Name = "Nhà phân phối đồ sưu tập, quà lưu niệm",
                    Phone = "21122334455",
                    Address = "Quận 1",
                    City = "Hồ Chí Minh",
                    Country = "Việt Nam",
                    Description = "Chuyên phân phối đồ sưu tập, quà lưu niệm",
                    Url = "supplier3.com"
                }
                );
            #endregion
        }
    }
}
