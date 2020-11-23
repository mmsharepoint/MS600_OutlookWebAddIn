using System;
using System.Configuration;
using System.Threading.Tasks;
using Microsoft.IdentityModel.Tokens;
using Microsoft.Owin;
using Microsoft.Owin.Security.Jwt;
using Microsoft.Owin.Security.OAuth;
using Owin;

[assembly: OwinStartup(typeof(OutlookWebAddInWeb.Startup))]

namespace OutlookWebAddInWeb
{
  public class Startup
  {
    public void Configuration(IAppBuilder app)
    {
      // For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=316888
      ConfigureAuth(app);
    }
    public void ConfigureAuth(IAppBuilder app)
    {
      string authority = String.Format(ConfigurationManager.AppSettings["Authority"], ConfigurationManager.AppSettings["DirectoryID"]);
      // string authority = String.Format(ConfigurationManager.AppSettings["Authority"], "common");
      TokenValidationParameters tvps = new TokenValidationParameters
      {
        ValidAudience = ConfigurationManager.AppSettings["AppIDUri"],
        ValidateIssuer = false,
        SaveSigninToken = true // Necessary to later have a BootstrapContext, that is the raw bootstraptoken
      };

      string[] endAuthoritySegments = { "oauth2/v2.0" };
      string[] parsedAuthority = authority.Split(endAuthoritySegments, System.StringSplitOptions.None);
      string wellKnownURL = parsedAuthority[0] + "v2.0/.well-known/openid-configuration";

      app.UseOAuthBearerAuthentication(new OAuthBearerAuthenticationOptions
      {
        AccessTokenFormat = new JwtFormat(tvps, new OpenIdConnectCachingSecurityTokenProvider(wellKnownURL))
      });
    }
  }
}
