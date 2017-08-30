using System;

namespace SpTaxonomyApiTester
{
    /// <summary>
    ///     Encapsulates all the information from SharePoint in ACS mode.
    /// </summary>
    internal class SharePointAcsContext : SharePointContext
    {
        private readonly string contextToken;
        private readonly SharePointContextToken contextTokenObj;

        public SharePointAcsContext(Uri spHostUrl,
            Uri spAppWebUrl,
            string spLanguage,
            string spClientTag,
            string spProductNumber,
            string contextToken,
            SharePointContextToken contextTokenObj)
            : base(spHostUrl, spAppWebUrl, spLanguage, spClientTag, spProductNumber)
        {
            if (string.IsNullOrEmpty(contextToken))
                throw new ArgumentNullException("contextToken");

            if (contextTokenObj == null)
                throw new ArgumentNullException("contextTokenObj");

            this.contextToken = contextToken;
            this.contextTokenObj = contextTokenObj;
        }

        /// <summary>
        ///     The context token's "CacheKey" claim.
        /// </summary>
        public string CacheKey
        {
            get { return contextTokenObj.ValidTo > DateTime.UtcNow ? contextTokenObj.CacheKey : null; }
        }

        /// <summary>
        ///     The context token.
        /// </summary>
        public string ContextToken
        {
            get { return contextTokenObj.ValidTo > DateTime.UtcNow ? contextToken : null; }
        }
    }
}