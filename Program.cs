using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;

class Program 
{

    static async Task Main(string[] args)
    {
        // azure app registrations info
        string clientId = "####";
        string clientSecret = "####";
        string tenantId = "####";
        string tokenEndpoint = $"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token";

        // create an httpclient
        using (HttpClient client = new HttpClient()) 
        {
            var requestContent = new FormUrlEncodedContent(new[]
            {
                new KeyValuePair<string, string>("grant_type", "client_credentials"),
                new KeyValuePair<string, string>("client_id", clientId),
                new KeyValuePair<string, string>("client_secret", clientSecret),
                new KeyValuePair<string, string>("scope", "https://graph.microsoft.com/.default"),

            });

            // retrieve access token
            HttpResponseMessage response = await client.PostAsync(tokenEndpoint, requestContent);
            string responseContent = await response.Content.ReadAsStringAsync();
            var tokenResponse = JsonSerializer.Deserialize<JsonElement>(responseContent);
            string accessToken = tokenResponse.GetProperty("access_token").GetString();
            // Console.WriteLine(accessToken);

            string mailUser = "";
            string sendMailEndpoint = $"https://graph.microsoft.com/v1.0/users/{mailUser}/sendMail";

            var message = new Dictionary<string, object>()
            {
                {"message", new Dictionary<string, object>()
                {
                    {"subject", "Test email using graph api, subject 2"},
                    {"body", new Dictionary<string, object>()
                    {
                        {"contentType", "Text"},
                        {"content", "Hello this is a successful message send by microsoft graph api with dot framework, message two"}
                    }
                    },
                    {"toRecipients", new object[]{
                        new Dictionary<string, object>()
                        {
                            {"emailAddress", new Dictionary<string, object>()
                            {
                                {"address", ""}
                            }
                            }
                        }
                    }},
                }

                },
                 {"saveToSentItems", "true"}
            };


            var jsonMessage = JsonSerializer.Serialize(message);
            var content = new StringContent(jsonMessage, Encoding.UTF8, "application/json");
            // post request to send the email
            var request = new HttpRequestMessage(HttpMethod.Post, sendMailEndpoint);
            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
            request.Content = content;

            HttpResponseMessage sendMailResponse = await client.SendAsync(request);
            string sendMailResponseContent = await sendMailResponse.Content.ReadAsStringAsync();

            if (sendMailResponse.IsSuccessStatusCode){
                Console.WriteLine("Email send successfully");
            }else{

                Console.WriteLine("Failed to send the email");
            }
        

        }

    }
}
