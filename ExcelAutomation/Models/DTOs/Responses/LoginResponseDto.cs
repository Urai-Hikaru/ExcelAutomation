namespace ExcelAutomation.Models.DTOs.Responses
{
    public class LoginResponseDto
    {
        public string Token { get; set; } = string.Empty;
        public int UserId { get; set; }
        public string Username { get; set; } = string.Empty;
        public string Rolename { get; set; } = string.Empty;
        public DateTime Expiration { get; set; }
    }
}