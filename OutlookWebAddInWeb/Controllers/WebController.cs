using Microsoft.Graph;
using Microsoft.Identity.Client;
using OutlookWebAddInWeb.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web.Http;

namespace OutlookWebAddInWeb.Controllers
{
  [Authorize]
  public class WebController : ApiController
  {
    [HttpGet]
    public string[] FileTypes()
    {
      string[] filtypes = { "docx", "pptx", "xlsx" };
      return filtypes;
    }

    [HttpPost]
    public async Task<string> StoreMimeMessage([FromBody] MimeMail request)
    {
      string[] graphScopes = { "https://graph.microsoft.com/Mail.Read", "https://graph.microsoft.com/Files.ReadWrite" };
      string accessToken = await GetAccessToken(graphScopes);

      string mimeContent = await GetMime(accessToken, request.MessageID);
      return mimeContent;
    }
    private async Task<string> GetAccessToken(string[] graphScopes)
    {
      string bootstrapContext = ClaimsPrincipal.Current.Identities.First().BootstrapContext.ToString();
      UserAssertion userAssertion = new UserAssertion(bootstrapContext);

      string authority = String.Format(ConfigurationManager.AppSettings["Authority"], ConfigurationManager.AppSettings["DirectoryID"]);

      string appID = ConfigurationManager.AppSettings["ClientID"];
      string appSecret = ConfigurationManager.AppSettings["ClientSecret"];
      var cca = ConfidentialClientApplicationBuilder.Create(appID)
                                                     .WithRedirectUri("https://localhost:44397")
                                                     .WithClientSecret(appSecret)
                                                     .WithAuthority(authority)
                                                     .Build();
      AcquireTokenOnBehalfOfParameterBuilder parameterBuilder = null;
      AuthenticationResult authResult = null;
      try
      {
        parameterBuilder = cca.AcquireTokenOnBehalfOf(graphScopes, userAssertion);
        authResult = await parameterBuilder.ExecuteAsync();
        return authResult.AccessToken;
      }
      catch (MsalServiceException e)
      {
        return null;
      }
    }
    private async Task<string> GetMime(string accessToken, string mailID)
    {
      GraphServiceClient graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(
          async (requestMessage) =>
          {
            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
          }));

      var mimeContent = await graphClient.Me.Messages[mailID]
        .Content
        .Request()
        .GetAsync();
      string mimeContentStr = string.Empty;
      using (var Reader = new System.IO.StreamReader(mimeContent))
      {
        mimeContentStr = Reader.ReadToEnd();
      }
      return mimeContentStr;
    }
  }
}
