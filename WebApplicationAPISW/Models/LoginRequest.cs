using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;

namespace WebApplicationAPISW.Models
{
    public class LoginRequest
    {
        [Display(Name = "Usuario")]
        public string Username { get; set; }
        [Display(Name = "Constraseña")]
        public string Password { get; set; }
    }
}