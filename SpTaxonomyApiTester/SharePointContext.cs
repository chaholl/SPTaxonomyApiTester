using System;
using System.Web;

namespace SpTaxonomyApiTester
{
    /// <summary>
    ///     Encapsulates all the information from SharePoint.
    /// </summary>
    internal abstract class SharePointContext
    {
        public const string SPHostUrlKey = "SPHostUrl";
        public const string SPAppWebUrlKey = "SPAppWebUrl";
        public const string SPLanguageKey = "SPLanguage";
        public const string SPClientTagKey = "SPClientTag";
        public const string SPProductNumberKey = "SPProductNumber";
        protected static readonly TimeSpan AccessTokenLifetimeTolerance = TimeSpan.FromMinutes(5.0);
        // <AccessTokenString, UtcExpiresOn>
        protected Tuple<string, DateTime> appOnlyAccessTokenForSPAppWeb;
        protected Tuple<string, DateTime> appOnlyAccessTokenForSPHost;
        protected Tuple<string, DateTime> userAccessTokenForSPAppWeb;
        protected Tuple<string, DateTime> userAccessTokenForSPHost;

        /// <summary>
        ///     Constructor.
        /// </summary>
        /// <param name="spHostUrl">The SharePoint host url.</param>
        /// <param name="spAppWebUrl">The SharePoint app web url.</param>
        /// <param name="spLanguage">The SharePoint language.</param>
        /// <param name="spClientTag">The SharePoint client tag.</param>
        /// <param name="spProductNumber">The SharePoint product number.</param>
        protected SharePointContext(Uri spHostUrl, Uri spAppWebUrl, string spLanguage, string spClientTag, string spProductNumber)
        {
            if (spHostUrl == null)
                throw new ArgumentNullException("spHostUrl");

            if (string.IsNullOrEmpty(spLanguage))
                throw new ArgumentNullException("spLanguage");

            if (string.IsNullOrEmpty(spClientTag))
                throw new ArgumentNullException("spClientTag");

            if (string.IsNullOrEmpty(spProductNumber))
                throw new ArgumentNullException("spProductNumber");

            SPHostUrl = spHostUrl;
            SPAppWebUrl = spAppWebUrl;
            SPLanguage = spLanguage;
            SPClientTag = spClientTag;
            SPProductNumber = spProductNumber;
        }

        /// <summary>
        ///     The SharePoint app web url.
        /// </summary>
        public Uri SPAppWebUrl { get; }

        /// <summary>
        ///     The SharePoint client tag.
        /// </summary>
        public string SPClientTag { get; }

        /// <summary>
        ///     The SharePoint host url.
        /// </summary>
        public Uri SPHostUrl { get; }

        /// <summary>
        ///     The SharePoint language.
        /// </summary>
        public string SPLanguage { get; }

        /// <summary>
        ///     The SharePoint product number.
        /// </summary>
        public string SPProductNumber { get; }

        /// <summary>
        ///     Gets the SharePoint host url from QueryString of the specified HTTP request.
        /// </summary>
        /// <param name="httpRequest">The specified HTTP request.</param>
        /// <returns>The SharePoint host url. Returns <c>null</c> if the HTTP request doesn't contain the SharePoint host url.</returns>
        public static Uri GetSPHostUrl(HttpRequestBase httpRequest)
        {
            if (httpRequest == null)
                throw new ArgumentNullException("httpRequest");

            var spHostUrlString = TokenHelper.EnsureTrailingSlash(httpRequest.QueryString[SPHostUrlKey]);
            Uri spHostUrl;
            if (Uri.TryCreate(spHostUrlString, UriKind.Absolute, out spHostUrl) &&
                (spHostUrl.Scheme == Uri.UriSchemeHttp || spHostUrl.Scheme == Uri.UriSchemeHttps))
                return spHostUrl;

            return null;
        }

        ///// <summary>
        /////     Gets the SharePoint host url from QueryString of the specified HTTP request.
        ///// </summary>
        ///// <param name="httpRequest">The specified HTTP request.</param>
        ///// <returns>The SharePoint host url. Returns <c>null</c> if the HTTP request doesn't contain the SharePoint host url.</returns>
        //public static Uri GetSPHostUrl(HttpRequest httpRequest)
        //{
        //    return GetSPHostUrl(new HttpRequestWrapper(httpRequest));
        //}

        /// <summary>
        ///     Determines if the specified access token is valid.
        ///     It considers an access token as not valid if it is null, or it has expired.
        /// </summary>
        /// <param name="accessToken">The access token to verify.</param>
        /// <returns>True if the access token is valid.</returns>
        protected static bool IsAccessTokenValid(Tuple<string, DateTime> accessToken)
        {
            return accessToken != null &&
                   !string.IsNullOrEmpty(accessToken.Item1) &&
                   accessToken.Item2 > DateTime.UtcNow;
        }
    }
      
}