using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CoffeeBook.Dto
{
    public class ForgotPassDto
    {
        private string newPassword;

        public string NewPassword { get => newPassword; set => newPassword = value; }
    }
}
