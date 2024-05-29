using System;
using System.Threading.Tasks;
using Microsoft.Exchange.WebServices.Data;

namespace EWSSyncExample {
    public static class Authorization {
        public static async Task<ExchangeService> GetExchangeWebService() {
            // Configure the MSAL client to get tokens            
            var pcaOptions = new Microsoft.Identity.Client.PublicClientApplicationOptions
            {
                ClientId = System.Configuration.ConfigurationManager.AppSettings["clientId"],
                TenantId = System.Configuration.ConfigurationManager.AppSettings["tenantId"],
                RedirectUri = "http://localhost"
            };

            var pca = Microsoft.Identity.Client.PublicClientApplicationBuilder
                .CreateWithApplicationOptions(pcaOptions).Build();

            // The permission scope required for EWS access
            var ewsScopes = new string[] { "https://outlook.office365.com/EWS.AccessAsUser.All" };

            // Make the interactive token request
            var authResult = await pca.AcquireTokenInteractive(ewsScopes).ExecuteAsync();

            // Configure the ExchangeService with the access token
            var ewsClient = new ExchangeService();
            ewsClient.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
            ewsClient.Credentials = new OAuthCredentials(authResult.AccessToken);
            return ewsClient;
        }
    }
}
