using System.Net.Http.Headers;
using Newtonsoft.Json;

namespace ProgKidsNotifier.Services;

public class MessageService
{
    private const string _channelIdTechSupp = "";
    private const string _channelIdTechNotifications = "";
    private const string _postUrl = "/api/v4/posts";
    private const string _botApiToken = "";
    
    private async Task<string?> SendToMattermost(string messageToSend)
    {
        var client = HttpClients.Default;
        var jsonPayload = new
        {
            message = messageToSend,
            channel_id = _channelIdTechNotifications
        };
        var content = new StringContent(
            Newtonsoft.Json.JsonConvert.SerializeObject(jsonPayload),
            System.Text.Encoding.UTF8,
            "application/json");
        
        var response = await client.PostAsync(_postUrl,  content);
        
        if (response.IsSuccessStatusCode)
        {
            var postId = await System.Text.Json.JsonSerializer.DeserializeAsync<MonitorService.MessageRespose>(
                await response.Content.ReadAsStreamAsync());
          //  var postId = JsonConvert.DeserializeObject<MonitorService.MessageRespose>(await response.Content.ReadAsStringAsync());
            Console.WriteLine($"Message sent to Mattermost successfully. | PostId : {postId}" );
            return postId?.id;
        }
        Console.WriteLine($"Failed to send message to Mattermost. | Repose {await response.Content.ReadAsStringAsync()}");
        return null;
    }
    
    private static class HttpClients
    {
        public static readonly HttpClient Default = new HttpClient
        {
            // Optionally set default headers, timeouts, etc.\
            DefaultRequestHeaders = { Authorization = new AuthenticationHeaderValue("Bearer", _botApiToken) },
            Timeout = TimeSpan.FromSeconds(15)
        };
    }
    
}