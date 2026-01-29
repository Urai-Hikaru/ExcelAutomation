using ExcelAutomation.Models;
using ExcelAutomation.Models.DTOs.Requests;
using ExcelAutomation.Models.DTOs.Responses;

namespace ExcelAutomation.Services.Api
{
    public class AuthService : BaseApiService, IAuthService
    {
        public AuthService() : base()
        {
        }

        public async Task<ApiResult<LoginResponseDto>> LoginAsync(string username, string password)
        {
            var requestDto = new LoginRequestDto
            {
                Username = username,
                Password = password
            };

            return await SendPostAsync<LoginRequestDto, LoginResponseDto>("Auth/Login", requestDto);
        }
    }
}