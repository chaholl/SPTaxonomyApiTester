using System;
using System.Net;
using System.Web;
using Microsoft.IdentityModel.Tokens;

namespace SpTaxonomyApiTester
{
    /// <summary>
    ///     Default provider for SharePointAcsContext.
    /// </summary>
    internal class SharePointAcsContextProvider : SharePointContextProvider
    {
        private const string SPContextKey = "SPContext";
        private const string SPCacheKeyKey = "SPCacheKey";

        protected override SharePointContext CreateSharePointContext(Uri spHostUrl,
            Uri spAppWebUrl,
            string spLanguage,
            string spClientTag,
            string spProductNumber,
            HttpRequestBase httpRequest)
        {
            var contextTokenString = TokenHelper.GetContextTokenFromRequest(httpRequest);
            if (string.IsNullOrEmpty(contextTokenString))
                return null;

            SharePointContextToken contextToken = null;
            try
            {
                if (httpRequest.Url != null) contextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, httpRequest.Url.Authority);
            }
            catch (WebException)
            {
                return null;
            }
            //catch (SecurityTokenExpiredException)
            //{
            //    return null;
            //}
            catch (AudienceUriValidationFailedException)
            {
                return null;
            }

            return new SharePointAcsContext(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber, contextTokenString, contextToken);
        }

        protected override SharePointContext LoadSharePointContext(HttpContextBase httpContext)
        {
            if (httpContext.Session != null) return httpContext.Session[SPContextKey] as SharePointAcsContext;
            return null;
        }

        protected override void SaveSharePointContext(SharePointContext spContext, HttpContextBase httpContext)
        {
            var spAcsContext = spContext as SharePointAcsContext;

            if (spAcsContext != null)
            {
                var spCacheKeyCookie = new HttpCookie(SPCacheKeyKey)
                {
                    Value = spAcsContext.CacheKey,
                    Secure = true,
                    HttpOnly = true
                };

                httpContext.Response.AppendCookie(spCacheKeyCookie);
            }

            if (httpContext.Session != null) httpContext.Session[SPContextKey] = spAcsContext;
        }

        protected override bool ValidateSharePointContext(SharePointContext spContext, HttpContextBase httpContext)
        {
            var spAcsContext = spContext as SharePointAcsContext;

            if (spAcsContext != null)
            {
                var spHostUrl = SharePointContext.GetSPHostUrl(httpContext.Request);
                var contextToken = TokenHelper.GetContextTokenFromRequest(httpContext.Request);
                var spCacheKeyCookie = httpContext.Request.Cookies[SPCacheKeyKey];
                var spCacheKey = spCacheKeyCookie != null ? spCacheKeyCookie.Value : null;

                return spHostUrl == spAcsContext.SPHostUrl &&
                       !string.IsNullOrEmpty(spAcsContext.CacheKey) &&
                       spCacheKey == spAcsContext.CacheKey &&
                       !string.IsNullOrEmpty(spAcsContext.ContextToken) &&
                       (string.IsNullOrEmpty(contextToken) || contextToken == spAcsContext.ContextToken);
            }

            return false;
        }
    }
}