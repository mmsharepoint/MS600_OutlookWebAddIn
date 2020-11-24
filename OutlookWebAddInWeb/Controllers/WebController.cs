using Microsoft.Graph;
using Microsoft.Identity.Client;
using OutlookWebAddInWeb.Models;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
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

      Stream mimeContent = await GetMime(accessToken, request.MessageID);
      string webUrl = await uploadMail2OD(accessToken, mimeContent, "Testmail.eml");
      return webUrl;
    }


    [HttpPost]
    public async Task<IMessageAttachmentsCollectionPage> GetAttachments([FromBody] MimeMail request)
    {
      string[] graphScopes = { "https://graph.microsoft.com/Mail.Read", "https://graph.microsoft.com/Files.ReadWrite" };
      string accessToken = await GetAccessToken(graphScopes);
      GraphServiceClient graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(
          async (requestMessage) =>
          {
            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
          }));

      var attachments = await graphClient.Me.Messages[request.MessageID]
        .Attachments
        .Request()
        .GetAsync();

      return attachments;
    }

    [HttpPost]
    public async Task<string[]> SaveAttachments([FromBody] Models.AttachmentRequest request)
    {
      string[] graphScopes = { "https://graph.microsoft.com/Mail.Read", "https://graph.microsoft.com/Files.ReadWrite" };
      string accessToken = await GetAccessToken(graphScopes);
      GraphServiceClient graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(
          async (requestMessage) =>
          {
            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
          }));
      List<string> resultUrls = new List<string>();
      foreach (Models.Attachment file in request.Attachments)
      {
        Stream attachment = await GetAttachment(accessToken, request.MessageID, file.id);
        string url = await uploadMail2OD(accessToken, attachment, file.filename);
        resultUrls.Add(url);
      }
      return resultUrls.ToArray();
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
    private async Task<Stream> GetMime(string accessToken, string mailID)
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
      //string mimeContentStr = string.Empty;
      //using (var Reader = new System.IO.StreamReader(mimeContent))
      //{
      //  mimeContentStr = Reader.ReadToEnd();
      //}
      return mimeContent;
    }

    private async Task<string> uploadMail2OD(string accessToken, Stream stream, string filename)
    {
      GraphServiceClient graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(
          async (requestMessage) =>
          {
            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
          }));

      if (stream.Length < (4 * 1024 * 1024))
      {
        DriveItem uploadResult = await graphClient.Me
                                                      .Drive.Root
                                                      .ItemWithPath(filename)
                                                      .Content.Request()
                                                      .PutAsync<DriveItem>(stream);
        return uploadResult.WebUrl;
      }
      else
      {
        try
        { // This method supports files even greater 4MB
          DriveItem item = null;
          UploadSession session = await graphClient.Me.Drive.Root
              .ItemWithPath(filename).CreateUploadSession().Request().PostAsync();          
          int maxSizeChunk = 320 * 4 * 1024;
          ChunkedUploadProvider provider = new ChunkedUploadProvider(session, graphClient, stream, maxSizeChunk);
          var chunckRequests = provider.GetUploadChunkRequests();
          List<Exception> exceptions = new List<Exception>();
          foreach (UploadChunkRequest chunkReq in chunckRequests) //upload the chunks
          {
            var reslt = await provider.GetChunkRequestResponseAsync(chunkReq, exceptions);
            if (reslt.UploadSucceeded) item = reslt.ItemResponse; // Check that upload succeeded
          }
          return item != null ? item.WebUrl : null;
        }
        catch (ServiceException ex)
        {
          return null;
        }        
      }
    }

    private async Task<Stream> GetAttachment(string accessToken, string mailID, string attachmentID)
    {
      GraphServiceClient graphClient = new GraphServiceClient(new DelegateAuthenticationProvider(
          async (requestMessage) =>
          {
            requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", accessToken);
          }));
      // Workaround for missing .Content impl. on .Attachments["ID"].
      var attachmentRequest = graphClient.Me.Messages[mailID]
                        .Attachments[attachmentID];
      var fileRequest = new FileAttachmentRequestBuilder(attachmentRequest.RequestUrl, graphClient);
      var u = fileRequest.Content.Request().RequestUrl;
      Stream fileContent = await fileRequest.Content.Request().GetAsync();      
      return fileContent;
    }
  }
}
