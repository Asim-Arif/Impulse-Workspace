using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAccessLibrary.Models.ViewModels
{
    public class LoginViewModel
    {
        [Required(AllowEmptyStrings = false, ErrorMessage ="Please prvide User Name")]
        //[Required(AllowEmptyStrings = false, ErrorMessage = "Please Provide Emp Code.")]
        public string? UserName {  get; set; }

        [Required(AllowEmptyStrings = false, ErrorMessage = "Please Provide Password")]
        public string? Password { get; set; }
    }
}
