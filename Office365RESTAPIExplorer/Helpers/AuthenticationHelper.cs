// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OAuth;
using System;
using System.Linq;
using System.Threading.Tasks;
using Windows.Security.Authentication.Web;

namespace Office365RESTExplorerforSites
{
    /// <summary>
    /// Provides access tokens for the Office 365 resources.
    /// </summary>
    internal static class AuthenticationHelper
    {
        // Properties of the native client app
        // The ClientID is added as a resource in App.xaml when you register the app with 
        // Office 365. As a convenience, we load that value into a variable called ClientID. By doing this, 
        // whenever you register the app using another account, this variable will be in sync with whatever is in App.xaml.
        private static readonly string ClientID = App.Current.Resources["ida:ClientID"].ToString();
        private static Uri ReturnUri = WebAuthenticationBroker.GetCurrentApplicationCallbackUri();


        // Properties used for communicating with the Windows Azure AD tenant of choice
        // The AuthorizationUri is added as a resource in App.xaml when you regiter the app with 
        // Office 365. As a convenience, we load that value into a variable called CommonAuthority, adding Common to this Url to signify
        // multi-tenancy. By doing this, whenever you register the app using another account, this variable will be in sync with whatever is in App.xaml.
        private static readonly string CommonAuthority = App.Current.Resources["ida:AuthorizationUri"].ToString() + @"/Common";

        public static AuthenticationContext AuthenticationContext { get; set; }

        /// <summary>
        /// Checks that an access token is available. 
        /// </summary>
        /// <returns>The Authentication result. </returns>
        public static async Task<AuthenticationResult> EnsureAccessTokenAvailableAsync(string serviceResourceId, PromptBehavior behavior)
        {
            AuthenticationContext = new AuthenticationContext(CommonAuthority);

            if (AuthenticationContext.TokenCache.ReadItems().Count() > 0)
            {
                // re-bind the AuthenticationContext to the authority that sourced the token in the cache 
                // this is needed for the cache to work when asking for a token from that authority 
                // (the common endpoint never triggers cache hits) 
                string cachedAuthority = AuthenticationContext.TokenCache.ReadItems().First().Authority;
                AuthenticationContext = new AuthenticationContext(cachedAuthority);

            }

            //Get the result of the of the token acquisition operation
            AuthenticationResult result = await AuthenticationContext.AcquireTokenAsync(serviceResourceId, ClientID, ReturnUri, behavior);

            return result;
        }
    }
}

//********************************************************* 
// 
//Office-365-REST-API-Explorer, https://github.com/OfficeDev/Office-365-REST-API-Explorer
//
//Copyright (c) Microsoft Corporation
//All rights reserved. 
//
//MIT License:
//
//Permission is hereby granted, free of charge, to any person obtaining
//a copy of this software and associated documentation files (the
//""Software""), to deal in the Software without restriction, including
//without limitation the rights to use, copy, modify, merge, publish,
//distribute, sublicense, and/or sell copies of the Software, and to
//permit persons to whom the Software is furnished to do so, subject to
//the following conditions:
//
//The above copyright notice and this permission notice shall be
//included in all copies or substantial portions of the Software.
//
//THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
//EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
//MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
//NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
//LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
//OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
//WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
// 
//********************************************************* 