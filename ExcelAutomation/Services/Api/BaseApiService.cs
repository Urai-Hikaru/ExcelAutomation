using System.Configuration;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using ExcelAutomation.Models;
using ExcelAutomation.Services.Common;

namespace ExcelAutomation.Services.Api
{
    public abstract class BaseApiService
    {
        protected readonly HttpClient _httpClient;
        protected readonly JsonSerializerOptions _jsonOptions;
        protected readonly SystemLogService _logger = SystemLogService.Instance;

        protected BaseApiService()
        {
            var handler = new HttpClientHandler
            {
                ServerCertificateCustomValidationCallback = (message, cert, chain, errors) => true
            };
            _httpClient = new HttpClient(handler);

            string? baseUrl = ConfigurationManager.AppSettings["BaseUrl"];
            if (string.IsNullOrEmpty(baseUrl))
            {
                var ex = new InvalidOperationException("App.config に 'BaseUrl' が設定されていません。");
                _logger.LogError(ex, "初期化エラー");
                throw ex;
            }

            _httpClient.BaseAddress = new Uri(baseUrl.Trim());
            _httpClient.Timeout = TimeSpan.FromSeconds(30);

            _jsonOptions = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            };
        }

        protected async Task<ApiResult<TResponse>> SendPostAsync<TRequest, TResponse>(string endpoint, TRequest requestData)
        {
            try
            {
                // リクエスト開始ログ
                _logger.LogInfo($"API Request Start: {endpoint}");

                var json = JsonSerializer.Serialize(requestData);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                // 通信実行
                var response = await _httpClient.PostAsync(endpoint, content);

                if (response.IsSuccessStatusCode)
                {
                    var jsonString = await response.Content.ReadAsStringAsync();
                    var data = JsonSerializer.Deserialize<TResponse>(jsonString, _jsonOptions);

                    // 成功ログ
                    _logger.LogInfo($"API Request Success: {endpoint}");

                    return ApiResult<TResponse>.Success(data!);
                }
                else
                {
                    // HTTPステータスコードによるエラー分岐
                    string msg = response.StatusCode switch
                    {
                        System.Net.HttpStatusCode.Unauthorized => "認証に失敗しました。IDまたはパスワードを確認してください。",
                        System.Net.HttpStatusCode.Forbidden => "アクセス権限がありません。",
                        System.Net.HttpStatusCode.NotFound => "接続先が見つかりません (404)。",
                        System.Net.HttpStatusCode.InternalServerError => "サーバー内部エラーが発生しました (500)。",
                        System.Net.HttpStatusCode.BadRequest => "不正なリクエストです (400)。",
                        _ => $"サーバーエラーが発生しました。(Code: {response.StatusCode})"
                    };

                    // 失敗ログ (Warningレベル)
                    _logger.LogWarning($"API Request Failed: {endpoint} | Status: {response.StatusCode} | Message: {msg}");

                    return ApiResult<TResponse>.Failure(msg);
                }
            }
            // 通信レベルのエラー分岐
            catch (HttpRequestException ex)
            {
                _logger.LogError(ex, $"Network Error: {endpoint}");
                return ApiResult<TResponse>.Failure("サーバーに接続できません。ネットワークを確認してください。");
            }
            catch (TaskCanceledException ex)
            {
                _logger.LogError(ex, $"Timeout Error: {endpoint}");
                return ApiResult<TResponse>.Failure("通信がタイムアウトしました。しばらく待ってから再試行してください。");
            }
            catch (JsonException ex)
            {
                _logger.LogError(ex, $"JSON Parse Error: {endpoint}");
                return ApiResult<TResponse>.Failure("サーバーからの応答が不正です。");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Unexpected Error: {endpoint}");
                return ApiResult<TResponse>.Failure($"予期せぬエラーが発生しました: {ex.Message}");
            }
        }
    }
}