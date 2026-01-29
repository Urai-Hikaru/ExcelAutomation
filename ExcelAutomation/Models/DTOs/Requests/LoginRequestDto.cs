using System.ComponentModel.DataAnnotations;

namespace ExcelAutomation.Models.DTOs.Requests
{
    public class LoginRequestDto
    {
        [Required]
        public string Username { get; set; } = string.Empty;

        [Required]
        public string Password { get; set; } = string.Empty;
    }
}