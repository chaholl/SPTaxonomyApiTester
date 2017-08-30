using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Diagnostics;
using System.Globalization;
using System.IdentityModel.Selectors;
using System.IdentityModel.Tokens;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Web;
using System.Web.Caching;
using System.Web.Script.Serialization;
using Microsoft.IdentityModel.S2S.Protocols.OAuth2;
using Microsoft.IdentityModel.S2S.Tokens;
using Microsoft.IdentityModel.SecurityTokenService;
using Microsoft.SharePoint.Client;
using AudienceRestriction = Microsoft.IdentityModel.Tokens.AudienceRestriction;
using AudienceUriValidationFailedException = Microsoft.IdentityModel.Tokens.AudienceUriValidationFailedException;
using SecurityTokenHandlerConfiguration = Microsoft.IdentityModel.Tokens.SecurityTokenHandlerConfiguration;
using X509SigningCredentials = Microsoft.IdentityModel.SecurityTokenService.X509SigningCredentials;

namespace SpTaxonomyApiTester
{
    internal static class TokenHelper
    {
        private const int MAX_RETRIES = 5;
        private const int RETRY_INTERVAL_MAX = 1000;
        private const int RETRY_INTERVAL_MIN = 10;

        // This class is used to get MetaData document from the global STS endpoint. It contains
        // methods to parse the MetaData document and get endpoints and STS certificate.
        public static class AcsMetadataParser
        {
            public static X509Certificate2 GetAcsSigningCert(string realm)
            {
                var document = GetMetadataDocument(realm);

                if (null != document.keys && document.keys.Count > 0)
                {
                    var signingKey = document.keys[0];

                    if (null != signingKey && null != signingKey.keyValue)
                        return new X509Certificate2(Encoding.UTF8.GetBytes(signingKey.keyValue.value));
                }

                throw new Exception("Metadata document does not contain ACS signing certificate.");
            }

            public static string GetDelegationServiceUrl(string realm)
            {
                var document = GetMetadataDocument(realm);

                var delegationEndpoint = document.endpoints.SingleOrDefault(e => e.protocol == DelegationIssuance);

                if (null != delegationEndpoint)
                    return delegationEndpoint.location;
                throw new Exception("Metadata document does not contain Delegation Service endpoint Url");
            }

            public static string GetStsUrl(string realm)
            {
                var document = GetMetadataDocument(realm);

                var s2sEndpoint = document.endpoints.SingleOrDefault(e => e.protocol == S2SProtocol);

                if (null != s2sEndpoint)
                    return s2sEndpoint.location;

                throw new Exception("Metadata document does not contain STS endpoint url");
            }

            private static JsonMetadataDocument GetMetadataDocument(string realm)
            {
                var acsMetadataEndpointUrlWithRealm = string.Format(CultureInfo.InvariantCulture,
                    "{0}?realm={1}",
                    GetAcsMetadataEndpointUrl(),
                    realm);
                byte[] acsMetadata;
                using (var webClient = new WebClient())
                {
                    acsMetadata = webClient.DownloadData(acsMetadataEndpointUrlWithRealm);
                }
                var jsonResponseString = Encoding.UTF8.GetString(acsMetadata);

                var serializer = new JavaScriptSerializer();
                var document = serializer.Deserialize<JsonMetadataDocument>(jsonResponseString);

                if (null == document)
                    throw new Exception("No metadata document found at the global endpoint " + acsMetadataEndpointUrlWithRealm);

                return document;
            }

            #region Nested type: JsonEndpoint

            private class JsonEndpoint
            {
                public string location { get; set; }
                public string protocol { get; set; }
                public string usage { get; set; }
            }

            #endregion

            #region Nested type: JsonKey

            private class JsonKey
            {
                public JsonKeyValue keyValue { get; set; }
                public string usage { get; set; }
            }

            #endregion

            #region Nested type: JsonKeyValue

            private class JsonKeyValue
            {
                public string type { get; set; }
                public string value { get; set; }
            }

            #endregion

            #region Nested type: JsonMetadataDocument

            private class JsonMetadataDocument
            {
                public List<JsonEndpoint> endpoints { get; set; }
                public List<JsonKey> keys { get; set; }
                public string serviceName { get; set; }
            }

            #endregion
        }

        /// <summary>
        ///     SharePoint principal.
        /// </summary>
        public const string SHARE_POINT_PRINCIPAL = "00000003-0000-0ff1-ce00-000000000000";

        /// <summary>
        ///     Lifetime of HighTrust access token, 12 hours.
        /// </summary>
        public static readonly TimeSpan HighTrustAccessTokenLifetime = TimeSpan.FromHours(12.0);

        #region public methods

        /// <summary>
        ///     Ensures that the specified URL ends with '/' if it is not null or empty.
        /// </summary>
        /// <param name="url">The url.</param>
        /// <returns>The url ending with '/' if it is not null or empty.</returns>
        public static string EnsureTrailingSlash(string url)
        {
            if (!string.IsNullOrEmpty(url) && url[url.Length - 1] != '/')
                return url + "/";

            return url;
        }
      /// <summary>
        ///     Retrieves an app-only access token from ACS to call the specified principal
        ///     at the specified targetHost. The targetHost must be registered for target principal.  If specified realm is
        ///     null, the "Realm" setting in web.config will be used instead.
        /// </summary>
        /// <param name="targetPrincipalName">Name of the target principal to retrieve an access token for</param>
        /// <param name="targetHost">Url authority of the target principal</param>
        /// <param name="targetRealm">Realm to use for the access token's nameid and audience</param>
        /// <param name="logger">A logger to capture details of the process</param>
        /// <returns>An access token with an audience of the target principal</returns>
        public static OAuth2AccessTokenResponse GetAppOnlyAccessToken(
            string targetPrincipalName,
            string targetHost,
            string targetRealm)
        {
            if (targetRealm == null)
                targetRealm = Realm;

            var resource = GetFormattedPrincipal(targetPrincipalName, targetHost, targetRealm);
            var clientId = GetFormattedPrincipal(ClientId, HostedAppHostName, targetRealm);

            if (string.IsNullOrWhiteSpace(ClientSecret))
                throw new InvalidDataException("ClientSecret is not set");
            var oauth2Request = OAuth2MessageFactory.CreateAccessTokenRequestWithClientCredentials(clientId, ClientSecret, resource);
            oauth2Request.Resource = resource;

            OAuth2AccessTokenResponse oauth2Response = null;

            Console.WriteLine("Creating access token request with clientId:{0}, clientSecret starting with: {1}, resource:{2}",
                ClientId,
                ClientSecret.Substring(0, 5),
                resource);

            var retrycount = 0;
            var rnd = new Random();

            while (retrycount < MAX_RETRIES)
            {
                // Get token
                var client = new OAuth2S2SClient();

                try
                {
                    oauth2Response = client.Issue(AcsMetadataParser.GetStsUrl(targetRealm), oauth2Request) as OAuth2AccessTokenResponse;
                    break;
                }
                catch (RequestFailedException e)
                {
                    var iex = e.InnerException as WebException;
                    if (iex != null && iex.Status != WebExceptionStatus.Timeout)
                        throw;
                }
                catch (WebException wex)
                {
                    if (wex.Response != null)
                    {
                        var stream = wex.Response.GetResponseStream();
                        if (stream != null)
                            using (var sr = new StreamReader(stream))
                            {
                                var responseText = sr.ReadToEnd();
                                throw new WebException(wex.Message + " - " + responseText, wex);
                            }
                    }
                    throw new WebException(wex.Message, wex);
                }
                //Wait a random amount of time
                Thread.Sleep(rnd.Next(RETRY_INTERVAL_MIN, RETRY_INTERVAL_MAX));
                retrycount++;
                Console.WriteLine("OAuth2 token issue timed out. Retrying {0} of {1}", retrycount, MAX_RETRIES);
            }

            if (oauth2Response == null)
                throw new Exception("Unable to authenticate with SharePoint. Please try again later");

            return oauth2Response;
        }

        /// <summary>
        ///     Uses the specified access token to create a client context
        /// </summary>
        /// <param name="targetUrl">Url of the target SharePoint site</param>
        /// <param name="accessToken">Access token to be used when calling the specified targetUrl</param>
        /// <returns>A ClientContext ready to call targetUrl with the specified access token</returns>
        public static ClientContext GetClientContextWithAccessToken(string targetUrl, string accessToken)
        {
            var clientContext = new ClientContext(targetUrl)
            {
                AuthenticationMode = ClientAuthenticationMode.Anonymous,
                FormDigestHandlingEnabled = false
            };

            clientContext.ExecutingWebRequest +=
                delegate(object oSender, WebRequestEventArgs webRequestEventArgs)
                {
                    webRequestEventArgs.WebRequestExecutor.RequestHeaders["Authorization"] =
                        "Bearer " + accessToken;
                };

            return clientContext;
        }

        public static ClientContext GetAppOnlyClientContextForUrl(string spHostUrl)
        {
            try
            {
                var targetWeb = new Uri(spHostUrl);
                var targetRealm = GetRealmFromTargetUrl(targetWeb);
                OAuth2AccessTokenResponse responseToken = null;
                var hasCache = HttpContext.Current != null && HttpContext.Current.Cache != null;
                string cacheKey = $"SemaphoreAppOnlyTokenfor:{spHostUrl}";
                if (hasCache)
                    responseToken = HttpContext.Current.Cache[cacheKey] as OAuth2AccessTokenResponse;
                if (responseToken == null)
                {
                    responseToken = GetAppOnlyAccessToken(SHARE_POINT_PRINCIPAL, targetWeb.Authority, targetRealm);
                    if (hasCache)
                    {
                        HttpContext.Current.Cache.Add(cacheKey, responseToken, null, responseToken.ExpiresOn, Cache.NoSlidingExpiration, CacheItemPriority.Default, null);
                        Trace.WriteLine("TokenHelper:Adding app-only token to cache");
                    }
                }
                else
                {
                    Trace.WriteLine("TokenHelper:Using cached app-only token");
                }
                return GetClientContextWithAccessToken(targetWeb.ToString(), responseToken.AccessToken);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to get AppOnlyClientContect - {0}", ex.Message);
                return null;
            }
        }
      /// <summary>
        ///     Retrieves the context token string from the specified request by looking for well-known parameter names in the
        ///     POSTed form parameters and the querystring. Returns null if no context token is found.
        /// </summary>
        /// <param name="request">HttpRequest in which to look for a context token</param>
        /// <returns>The context token string</returns>
        public static string GetContextTokenFromRequest(HttpRequestBase request)
        {
            string[] paramNames = {"AppContext", "AppContextToken", "AccessToken", "SPAppToken"};
            foreach (var paramName in paramNames)
            {
                if (!string.IsNullOrEmpty(request.Form[paramName]))
                    return request.Form[paramName];
                if (!string.IsNullOrEmpty(request.QueryString[paramName]))
                    return request.QueryString[paramName];
            }
            return null;
        }

        /// <summary>
        ///     Get authentication realm from SharePoint
        /// </summary>
        /// <param name="targetApplicationUri">Url of the target SharePoint site</param>
        /// <returns>String representation of the realm GUID</returns>
        public static string GetRealmFromTargetUrl(Uri targetApplicationUri)
        {
            var request = WebRequest.Create(targetApplicationUri + "/_vti_bin/client.svc");
            request.Headers.Add("Authorization: Bearer ");

            try
            {
                using (request.GetResponse())
                {
                }
            }
            catch (WebException e)
            {
                if (e.Response == null)
                    return null;

                var bearerResponseHeader = e.Response.Headers["WWW-Authenticate"];
                if (string.IsNullOrEmpty(bearerResponseHeader))
                    return null;

                const string bearer = "Bearer realm=\"";
                var bearerIndex = bearerResponseHeader.IndexOf(bearer, StringComparison.Ordinal);
                if (bearerIndex < 0)
                    return null;

                var realmIndex = bearerIndex + bearer.Length;

                if (bearerResponseHeader.Length >= realmIndex + 36)
                {
                    var targetRealm = bearerResponseHeader.Substring(realmIndex, 36);

                    Guid realmGuid;

                    if (Guid.TryParse(targetRealm, out realmGuid))
                        return targetRealm;
                }
            }
            return null;
        }

        /// <summary>
        ///     Validate that a specified context token string is intended for this application based on the parameters
        ///     specified in web.config. Parameters used from web.config used for validation include ClientId,
        ///     HostedAppHostNameOverride, HostedAppHostName, ClientSecret, and Realm (if it is specified). If
        ///     HostedAppHostNameOverride is present,
        ///     it will be used for validation. Otherwise, if the <paramref name="appHostName" /> is not
        ///     null, it is used for validation instead of the web.config's HostedAppHostName. If the token is invalid, an
        ///     exception is thrown. If the token is valid, TokenHelper's static STS metadata url is updated based on the token
        ///     contents
        ///     and a JsonWebSecurityToken based on the context token is returned.
        /// </summary>
        /// <param name="contextTokenString">The context token to validate</param>
        /// <param name="appHostName">
        ///     The URL authority, consisting of  Domain Name System (DNS) host name or IP address and the port number, to use for
        ///     token audience validation.
        ///     If null, HostedAppHostName web.config setting is used instead. HostedAppHostNameOverride web.config setting, if
        ///     present, will be used
        ///     for validation instead of <paramref name="appHostName" /> .
        /// </param>
        /// <returns>A JsonWebSecurityToken based on the context token.</returns>
        /// <exception cref="SecurityTokenExpiredException" />
        /// <exception cref="Microsoft.IdentityModel.Tokens.AudienceUriValidationFailedException" />
        public static SharePointContextToken ReadAndValidateContextToken(string contextTokenString, string appHostName = null)
        {
            var tokenHandler = CreateJsonWebSecurityTokenHandler();
            var securityToken = tokenHandler.ReadToken(contextTokenString);
            var jsonToken = securityToken as JsonWebSecurityToken;
            var token = SharePointContextToken.Create(jsonToken);

            var stsAuthority = new Uri(token.SecurityTokenServiceUri).Authority;
            var firstDot = stsAuthority.IndexOf('.');

            GlobalEndPointPrefix = stsAuthority.Substring(0, firstDot);
            AcsHostUrl = stsAuthority.Substring(firstDot + 1);

            tokenHandler.ValidateToken(jsonToken);

            string[] acceptableAudiences;
            if (!string.IsNullOrEmpty(HostedAppHostNameOverride))
                acceptableAudiences = HostedAppHostNameOverride.Split(';');
            else if (appHostName == null)
                acceptableAudiences = new[] {HostedAppHostName};
            else
                acceptableAudiences = new[] {appHostName};

            var validationSuccessful = false;
            var realm = Realm ?? token.Realm;
            foreach (var audience in acceptableAudiences)
            {
                var principal = GetFormattedPrincipal(ClientId, audience, realm);
                if (StringComparer.OrdinalIgnoreCase.Equals(token.Audience, principal))
                {
                    validationSuccessful = true;
                    break;
                }
            }

            if (!validationSuccessful)
                throw new AudienceUriValidationFailedException($"\"{string.Join(";", acceptableAudiences)}\" is not the intended audience \"{token.Audience}\"");

            return token;
        }

        #endregion

        #region private fields

        //
        // Configuration Constants
        //        

        private const string AuthorizationPage = "_layouts/15/OAuthAuthorize.aspx";
        private const string RedirectPage = "_layouts/15/AppRedirect.aspx";
        private const string AcsPrincipalName = "00000001-0000-0000-c000-000000000000";
        private const string AcsMetadataEndPointRelativeUrl = "metadata/json/1";
        private const string S2SProtocol = "OAuth2";
        private const string DelegationIssuance = "DelegationIssuance1.0";
        private const string NameIdentifierClaimType = JsonWebTokenConstants.ReservedClaims.NameIdentifier;
        private const string TrustedForImpersonationClaimType = "trustedfordelegation";
        private const string ActorTokenClaimType = JsonWebTokenConstants.ReservedClaims.ActorToken;

        //
        // Environment Constants
        //

        private static string GlobalEndPointPrefix = "accounts";
        private static string AcsHostUrl = "accesscontrol.windows.net";

        //
        // Hosted app configuration
        //
        public static readonly string ClientId = ConfigurationManager.AppSettings.Get("ClientId");
        private static readonly string IssuerId = ClientId;

        private static readonly string HostedAppHostNameOverride = "";
        private static readonly string HostedAppHostName = "";

        private static readonly string ClientSecret = ConfigurationManager.AppSettings.Get("ClientSecret");
        
        private static readonly string Realm = ""; 
        private static readonly string ServiceNamespace = ""; 
     
        #endregion

        #region private methods


        private static JsonWebSecurityTokenHandler CreateJsonWebSecurityTokenHandler()
        {
            if (string.IsNullOrWhiteSpace(ClientSecret))
                throw new ConfigurationErrorsException("ClientSecret is not configured. Please review config file");

            var handler = new JsonWebSecurityTokenHandler();
            handler.Configuration = new SecurityTokenHandlerConfiguration();
            handler.Configuration.AudienceRestriction = new AudienceRestriction(AudienceUriMode.Never);
            handler.Configuration.CertificateValidator = X509CertificateValidator.None;

            var securityKeys = new List<byte[]>();
            securityKeys.Add(Convert.FromBase64String(ClientSecret));
        

            var securityTokens = new List<SecurityToken>();
            securityTokens.Add(new MultipleSymmetricKeySecurityToken(securityKeys));

            handler.Configuration.IssuerTokenResolver =
                SecurityTokenResolver.CreateDefaultSecurityTokenResolver(
                    new ReadOnlyCollection<SecurityToken>(securityTokens),
                    false);
            var issuerNameRegistry = new SymmetricKeyIssuerNameRegistry();
            foreach (var securitykey in securityKeys)
                issuerNameRegistry.AddTrustedIssuer(securitykey, GetAcsPrincipalName(ServiceNamespace));
            handler.Configuration.IssuerNameRegistry = issuerNameRegistry;
            return handler;
        }

        private static string GetAcsGlobalEndpointUrl()
        {
            return string.Format(CultureInfo.InvariantCulture, "https://{0}.{1}/", GlobalEndPointPrefix, AcsHostUrl);
        }

        private static string GetAcsMetadataEndpointUrl()
        {
            return Path.Combine(GetAcsGlobalEndpointUrl(), AcsMetadataEndPointRelativeUrl);
        }

        private static string GetAcsPrincipalName(string realm)
        {
            return GetFormattedPrincipal(AcsPrincipalName, new Uri(GetAcsGlobalEndpointUrl()).Host, realm);
        }


        private static string GetFormattedPrincipal(string principalName, string hostName, string realm)
        {
            if (!string.IsNullOrEmpty(hostName))
                return string.Format(CultureInfo.InvariantCulture, "{0}/{1}@{2}", principalName, hostName, realm);

            return string.Format(CultureInfo.InvariantCulture, "{0}@{1}", principalName, realm);
        }

        #endregion
    }
}