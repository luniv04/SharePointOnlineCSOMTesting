using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using PnP.Framework;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;

var scopes = new[] { "https://microsoft.sharepoint-df.com/.default" };
var tenantId = "6397730c-2970-4413-b48f-1abf00895d39";
var clientId = "826f164a-3c6c-4fdc-ba36-ed01442c1008";
var clientSecret = "";
var siteUrl = "https://78q4t7.sharepoint.com/sites/test1";

var clientApplication = ConfidentialClientApplicationBuilder.Create(clientId)
	.WithClientSecret(clientSecret)
	.WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
	.Build();

var result = await clientApplication.AcquireTokenForClient(scopes).ExecuteAsync();
var bearerToken = result.AccessToken;

Example0(siteUrl, clientId, clientSecret);
Example1(siteUrl, clientId, bearerToken);
await Example2(siteUrl, bearerToken);
await Example3(siteUrl, bearerToken);
await Example4(siteUrl, clientId, clientSecret, bearerToken, tenantId);
await Example5(siteUrl, bearerToken);
await Example6(siteUrl, clientId, clientSecret, tenantId);
await Example7(siteUrl, clientId, clientSecret, tenantId);
Example8(siteUrl, clientId, bearerToken);

// See https://answers.microsoft.com/en-us/msoffice/forum/all/sharepoint-app-only-add-ins-throwing-401/962bfaa2-8604-4e94-ae1c-36ef5b453ed2
// See https://www.sharepointdiary.com/2019/06/sharepoint-online-remote-server-returned-an-error-401-unauthorized.html#:~:text=Legacy%20authentication%20protocol%20is%20enabled%3F%20Check%20if%20the,enable%20it%20with%20the%20following%3A%20Set-SPOTenant%20-LegacyAuthProtocolsEnabled%20%24True
// "Get-PnPTenant | select LegacyAuthProtocolsEnabled" already is set to true.
// Tried using "Set-PnPTenant -DisableCustomAppAuthentication $false" because this setting was originally true but its not helping.

static void Example0(string siteUrl, string clientId, string clientSecret)
{
	try
	{
		using (var ctx = new AuthenticationManager().GetACSAppOnlyContext(siteUrl, clientId, clientSecret))
		{
			ctx.Load(ctx.Web, web => web.Title);
			ctx.ExecuteQuery();

			Console.ForegroundColor = ConsoleColor.Green;
			Console.WriteLine(ctx.Web.Title);
			Console.ResetColor();
		}
	}
	catch (Exception ex)
	{
		Console.ForegroundColor = ConsoleColor.Red;
		Console.WriteLine($"Example0 Failed: {ex.Message}");
		Console.ResetColor();
	}
}

static void Example1(string siteUrl, string clientId, string bearerToken)
{
	try
	{
		using (var ctx = new AuthenticationManager().GetACSAppOnlyContext(siteUrl, clientId, bearerToken))
		{
			ctx.Load(ctx.Web, web => web.Title);
			ctx.ExecuteQuery();

			Console.ForegroundColor = ConsoleColor.Green;
			Console.WriteLine(ctx.Web.Title);
			Console.ResetColor();
		}
	}
	catch(Exception ex)
	{
		Console.ForegroundColor = ConsoleColor.Red;
		Console.WriteLine($"Example1 Failed: {ex.Message}");
		Console.ResetColor();
	}
}

static async Task Example2(string siteUrl, string bearerToken)
{
	try
	{
		using (var client = new HttpClient())
		{
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
			client.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/json;odata=nometadata");

			var rawJson = await client.GetStringAsync($"{siteUrl}/_api/web?$select=Title");

			Console.ForegroundColor = ConsoleColor.Green;
			Console.WriteLine(rawJson);
			Console.ResetColor();
		}
	}
	catch (Exception ex)
	{
		Console.ForegroundColor = ConsoleColor.Red;
		Console.WriteLine($"Example2 Failed: {ex.Message}");
		Console.ResetColor();
	}
}

static async Task Example3(string siteUrl, string bearerToken)
{
	try
	{
		using (var ctx = new AuthenticationManager().GetAccessTokenContext(siteUrl, bearerToken))
		{
			var currentWeb = ctx.Web;
			ctx.Load(currentWeb);
			await ctx.ExecuteQueryRetryAsync();

			Console.ForegroundColor = ConsoleColor.Green;
			Console.WriteLine(currentWeb.Title);
			Console.ResetColor();
		}
	}
	catch (Exception ex)
	{
		Console.ForegroundColor = ConsoleColor.Red;
		Console.WriteLine($"Example3 Failed: {ex.Message}");
		Console.ResetColor();
	}
}

static async Task Example4(string siteUrl, string clientId, string clientSecret, string bearerToken, string tenantId)
{
	try
	{
		using (var ctx = new AuthenticationManager(clientId, clientSecret, new UserAssertion(bearerToken), tenantId).GetContext(siteUrl))
		{
			ctx.Load(ctx.Web, web => web.Title);
			await ctx.ExecuteQueryAsync();

			Console.ForegroundColor = ConsoleColor.Green;
			Console.WriteLine(ctx.Web.Title);
			Console.ResetColor();
		}
	}
	catch (Exception ex)
	{
		Console.ForegroundColor = ConsoleColor.Red;
		Console.WriteLine($"Example4 Failed: {ex.Message}");
		Console.ResetColor();
	}
}

async Task Example5(string siteUrl, string bearerToken)
{
	try
	{
		var url = $"{siteUrl}/_api/Web/Lists";
		var result = string.Empty;
		using (var client = new HttpClient())
		{
			client.DefaultRequestHeaders.Clear();
			client.DefaultRequestHeaders.Add("Accept", "application/json;odata=nometadata");
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
			var response = await client.GetAsync(url);
			response.EnsureSuccessStatusCode();
			result = await response.Content.ReadAsStringAsync();

			Console.ForegroundColor = ConsoleColor.Green;
			Console.WriteLine(result);
			Console.ResetColor();
		}
	}
	catch (Exception ex)
	{
		Console.ForegroundColor = ConsoleColor.Red;
		Console.WriteLine($"Example5 Failed: {ex.Message}");
		Console.ResetColor();
	}
}

async Task Example6(string siteUrl, string clientId, string clientSecret, string tenantId)
{
	var authorities = new List<string>
	{
		$"https://login.microsoftonline.com/{tenantId}",
		$"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/authorize",
		$"https://login.microsoftonline.com/{tenantId}/oauth2/v2.0/token",
		$"https://login.microsoftonline.com/{tenantId}/oauth2/authorize",
		$"https://login.microsoftonline.com/{tenantId}/oauth2/token",
		$"https://login.microsoftonline.com/{tenantId}/v2.0/.well-known/openid-configuration",
		"https://graph.microsoft.com",
		$"https://login.microsoftonline.com/{tenantId}/federationmetadata/2007-06/federationmetadata.xml",
		$"https://login.microsoftonline.com/{tenantId}/wsfed",
		$"https://login.microsoftonline.com/{tenantId}/saml2"
	};

	foreach (var authority in authorities)
	{
		try
		{
			var clientApplication = ConfidentialClientApplicationBuilder.Create(clientId)
				.WithClientSecret(clientSecret)
				.WithAuthority(new Uri(authority))
				.Build();

			var result = await clientApplication.AcquireTokenForClient(scopes).ExecuteAsync();
			var bearerToken = result.AccessToken;

			using (var client = new HttpClient())
			{
				client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
				client.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/json;odata=nometadata");

				var rawJson = await client.GetStringAsync($"{siteUrl}/_api/web?$select=Title");

				Console.ForegroundColor = ConsoleColor.Green;
				Console.WriteLine(rawJson);
				Console.ResetColor();
			}
		}
		catch (Exception ex)
		{
			Console.ForegroundColor = ConsoleColor.Red;
			Console.WriteLine($"Example6 Authority '{authority}' Failed: {ex.Message}");
			Console.ResetColor();
		}
	}
}

async Task Example7(string siteUrl, string clientId, string clientSecret, string tenantId)
{
	try
	{
		string certThumbPrint = "64c52dc51ef8cd01410752671d03dc872e05d66a";
		X509Store certStore = new X509Store(StoreName.My, StoreLocation.CurrentUser);
		// Try to open the store.

		certStore.Open(OpenFlags.ReadOnly);
		// Find the certificate that matches the thumbprint.
		X509Certificate2Collection certCollection = certStore.Certificates.Find(
			X509FindType.FindByThumbprint, certThumbPrint, false);
		certStore.Close();

		var firstCert = certCollection[0];

		var clientApplication = ConfidentialClientApplicationBuilder.Create(clientId)
			.WithCertificate(firstCert)
			.WithAuthority(new Uri($"https://login.microsoftonline.com/{tenantId}"))
			.Build();

		var result = await clientApplication.AcquireTokenForClient(scopes).ExecuteAsync();
		var bearerToken = result.AccessToken;

		using (var client = new HttpClient())
		{
			client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
			client.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/json;odata=nometadata");

			var rawJson = await client.GetStringAsync($"{siteUrl}/_api/web?$select=Title");

			Console.ForegroundColor = ConsoleColor.Green;
			Console.WriteLine(rawJson);
			Console.ResetColor();
		}
	}
	catch (Exception ex)
	{
		Console.ForegroundColor = ConsoleColor.Red;
		Console.WriteLine($"Example7 Failed: {ex.Message}");
		Console.ResetColor();
	}
}

static void Example8(string siteUrl, string clientId, string bearerToken)
{
	try
	{
		using (var ctx = new AuthenticationManager().GetACSAppOnlyContext(siteUrl, clientId, bearerToken))
		{
			ctx.ExecutingWebRequest += (sender, e) =>
			{
				e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + bearerToken;
			};
			ctx.Load(ctx.Web, web => web.Title);
			ctx.ExecuteQuery();

			Console.ForegroundColor = ConsoleColor.Green;
			Console.WriteLine(ctx.Web.Title);
			Console.ResetColor();
		}
	}
	catch (Exception ex)
	{
		Console.ForegroundColor = ConsoleColor.Red;
		Console.WriteLine($"Example8 Failed: {ex.Message}");
		Console.ResetColor();
	}
}
