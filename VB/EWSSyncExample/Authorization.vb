Imports System
Imports System.Threading.Tasks
Imports Microsoft.Exchange.WebServices.Data

Namespace EWSSyncExample
	Public Module Authorization
		Public Async Function GetExchangeWebService() As Task(Of ExchangeService)
			' Configure the MSAL client to get tokens            
			Dim pcaOptions = New Microsoft.Identity.Client.PublicClientApplicationOptions With {
				.ClientId = System.Configuration.ConfigurationManager.AppSettings("clientId"),
				.TenantId = System.Configuration.ConfigurationManager.AppSettings("tenantId"),
				.RedirectUri = "http://localhost"
			}

			Dim pca = Microsoft.Identity.Client.PublicClientApplicationBuilder.CreateWithApplicationOptions(pcaOptions).Build()

			' The permission scope required for EWS access
			Dim ewsScopes = New String() { "https://outlook.office365.com/EWS.AccessAsUser.All" }

			' Make the interactive token request
			Dim authResult = Await pca.AcquireTokenInteractive(ewsScopes).ExecuteAsync()

			' Configure the ExchangeService with the access token
			Dim ewsClient = New ExchangeService()
			ewsClient.Url = New Uri("https://outlook.office365.com/EWS/Exchange.asmx")
			ewsClient.Credentials = New OAuthCredentials(authResult.AccessToken)
			Return ewsClient
		End Function
	End Module
End Namespace
