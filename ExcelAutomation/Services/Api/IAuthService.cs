using ExcelAutomation.Models;
using ExcelAutomation.Models.DTOs.Responses;

namespace ExcelAutomation.Services.Api
{
    public interface IAuthService
    {
        Task<ApiResult<LoginResponseDto>> LoginAsync(string username, string password);
    }
}