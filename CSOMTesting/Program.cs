using Microsoft.Identity.Client;
using Microsoft.SharePoint.Client;
using PnP.Framework;
using System.Net.Http.Headers;

var scopes = new[] { "https://78q4t7.sharepoint.com/.default" };
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

//Example1(siteUrl, clientId, bearerToken);
await Example2(siteUrl, bearerToken);
await Example3(siteUrl, bearerToken);
await Example4(siteUrl, clientId, clientSecret, bearerToken, tenantId);
await Example5(siteUrl, bearerToken);


static void Example1(string siteUrl, string clientId, string bearerToken)
{
	using (var ctx = new AuthenticationManager().GetACSAppOnlyContext(siteUrl, clientId, bearerToken))
	{
		ctx.Load(ctx.Web, web => web.Title);
		ctx.ExecuteQuery();
		Console.WriteLine(ctx.Web.Title);
	}
}

static async Task Example2(string siteUrl, string bearerToken)
{
	using (var client = new HttpClient())
	{
		client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
		client.DefaultRequestHeaders.TryAddWithoutValidation("Accept", "application/json;odata=nometadata");

		var rawJson = await client.GetStringAsync($"{siteUrl}/_api/web?$select=Title");
		Console.WriteLine(rawJson);
	}
}

static async Task Example3(string siteUrl, string bearerToken)
{
	using (var ctx = new AuthenticationManager().GetAccessTokenContext(siteUrl, bearerToken))
	{
		var currentWeb = ctx.Web;
		ctx.Load(currentWeb);
		await ctx.ExecuteQueryRetryAsync();
		Console.WriteLine(currentWeb.Title);
	}
}

static async Task Example4(string siteUrl, string clientId, string clientSecret, string bearerToken, string tenantId)
{
	using (var ctx = new AuthenticationManager(clientId, clientSecret, new UserAssertion(bearerToken), tenantId).GetContext(siteUrl))
	{
		ctx.Load(ctx.Web, web => web.Title);
		await ctx.ExecuteQueryAsync();
		Console.WriteLine(ctx.Web.Title);
	}
}

async Task Example5(string siteUrl, string bearerToken)
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
	}
}