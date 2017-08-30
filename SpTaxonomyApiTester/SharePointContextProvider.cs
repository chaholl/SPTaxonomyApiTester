using System;
using System.Web;

namespace SpTaxonomyApiTester
{
    /// <summary>
    ///     Provides SharePointContext instances.
    /// </summary>
    abstract class SharePointContextProvider
    {
        /// <summary>
        ///     Initializes the default SharePointContextProvider instance.
        /// </summary>
        static SharePointContextProvider()
        {
            Current = new SharePointAcsContextProvider();
        }

        //end of custom code

        /// <summary>
        ///     The current SharePointContextProvider instance.
        /// </summary>
        public static SharePointContextProvider Current { get; private set; }

      
        /// <summary>
        ///     Creates a SharePointContext instance.
        /// </summary>
        /// <param name="spHostUrl">The SharePoint host url.</param>
        /// <param name="spAppWebUrl">The SharePoint app web url.</param>
        /// <param name="spLanguage">The SharePoint language.</param>
        /// <param name="spClientTag">The SharePoint client tag.</param>
        /// <param name="spProductNumber">The SharePoint product number.</param>
        /// <param name="httpRequest">The HTTP request.</param>
        /// <returns>The SharePointContext instance. Returns <c>null</c> if errors occur.</returns>
        protected abstract SharePointContext CreateSharePointContext(Uri spHostUrl,
            Uri spAppWebUrl,
            string spLanguage,
            string spClientTag,
            string spProductNumber,
            HttpRequestBase httpRequest);

        /// <summary>
        ///     Loads the SharePointContext instance associated with the specified HTTP context.
        /// </summary>
        /// <param name="httpContext">The HTTP context.</param>
        /// <returns>The SharePointContext instance. Returns <c>null</c> if not found.</returns>
        protected abstract SharePointContext LoadSharePointContext(HttpContextBase httpContext);


        /// <summary>
        ///     Saves the specified SharePointContext instance associated with the specified HTTP context.
        ///     <c>null</c> is accepted for clearing the SharePointContext instance associated with the HTTP context.
        /// </summary>
        /// <param name="spContext">The SharePointContext instance to be saved, or <c>null</c>.</param>
        /// <param name="httpContext">The HTTP context.</param>
        protected abstract void SaveSharePointContext(SharePointContext spContext, HttpContextBase httpContext);

        /// <summary>
        ///     Validates if the given SharePointContext can be used with the specified HTTP context.
        /// </summary>
        /// <param name="spContext">The SharePointContext.</param>
        /// <param name="httpContext">The HTTP context.</param>
        /// <returns>True if the given SharePointContext can be used with the specified HTTP context.</returns>
        protected abstract bool ValidateSharePointContext(SharePointContext spContext, HttpContextBase httpContext);
    }
}