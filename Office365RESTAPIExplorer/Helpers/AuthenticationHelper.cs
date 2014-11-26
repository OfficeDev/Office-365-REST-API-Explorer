// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.Discovery;
using Microsoft.Office365.OAuth;
using System;
using System.Linq;
using System.Threading.Tasks;
using Windows.Security.Authentication.Web;
using Office365RESTExplorerforSites.Data;

namespace Office365RESTExplorerforSites.Helpers
{
    /// <summary>
    /// Provides access tokens for the Office 365 resources.
    /// </summary>
    internal static class AuthenticationHelper
    {
        // Properties used for communicating with the Windows Azure AD tenant of choice
        // The AuthorizationUri is added as a resource in App.xaml when you regiter the app with 
        // Office 365. As a convenience, we load that value into a variable called CommonAuthority, adding Common to this Url to signify
        // multi-tenancy. By doing this, whenever you register the app using another account, this variable will be in sync with whatever is in App.xaml.
        private static readonly string _commonAuthority = App.Current.Resources["ida:AuthorizationUri"].ToString() + @"/Common";

        // Private properties used to store the access, refresh tokens
        // and the access token expiration date
        private static string _accessToken = null;
        private static DateTimeOffset _accessTokenExpiresOn = DateTimeOffset.MinValue.AddSeconds(10);
        private static string _refreshToken = null;

        private static AuthenticationContext _authenticationContext;

        // We need an event to notify other classes that the
        // access token has changed and they need to handle this.
        // For example, they need to update the data source
        public static event EventHandler AccessTokenChanged = delegate { };

        private static string _loggedInUser; 
        /// <summary> 
        /// Gets the logged in user. 
        /// </summary> 
        static internal String LoggedInUser 
        { 
            get 
            { 
                return _loggedInUser; 
            } 
        }

        private static string _serviceResourceId;
        /// <summary> 
        /// Gets the service resource id
        /// </summary> 
        static internal String ServiceResourceId
        {
            get
            {
                return _serviceResourceId;
            }
        }

        private static string _userAccount;
        /// <summary> 
        /// Gets the account of the user in the form user@domain.tld
        /// </summary> 
        static internal String UserAccount
        {
            get
            {
                return _userAccount;
            }
        }

        /// <summary>
        /// Checks that an access token is available.
        /// </summary>
        /// <returns>The access token.</returns>
        public static async Task<string> EnsureAccessTokenAvailableAsync()
        {
            if(!String.IsNullOrEmpty(_serviceResourceId))
            {
                return await EnsureAccessTokenAvailableAsync(_serviceResourceId);
            }
            else
            {
                MissingConfigurationValueException mcve = new MissingConfigurationValueException("To use this method you have to call EnsureAccessTokenCreatedAsync(string serviceResourceId) at least once.");
                MessageDialogHelper.DisplayException(mcve);
                return null;
            }
        }

        /// <summary>
        /// Checks that an access token is available.
        /// </summary>
        /// <returns>The access token.</returns>
        public static async Task<string> EnsureAccessTokenAvailableAsync(string serviceResourceId)
        {
            // If the token is not null nor empty and 
            // it it will not expire in the next 10 seconds
            bool tokenExpired = DateTimeOffset.Compare(_accessTokenExpiresOn.AddSeconds(-10), DateTimeOffset.Now) < 0;
            if (!String.IsNullOrEmpty(_accessToken) && !tokenExpired)
            {
                return _accessToken;
            }
            else
            {
                try
                {
                    string authority = _commonAuthority;

                    TokenCacheItem cacheItem = null;

                    // Create an AuthenticationContext using this authority.
                    _authenticationContext = new AuthenticationContext(authority, true);

                    //Get the current app object, which exposes the ClientId and ReturnUri properties
                    // that we need in the following call to AcquireTokenAsync
                    App currentApp = (App)App.Current;
                    
                    AuthenticationResult authenticationResult;
                    if (!String.IsNullOrEmpty(_refreshToken) && tokenExpired)
                    {
                        authenticationResult = await _authenticationContext.AcquireTokenByRefreshTokenAsync(_refreshToken, currentApp.ClientId, _serviceResourceId);
                    }
                    else
                    {
                        authenticationResult = await _authenticationContext.AcquireTokenAsync(serviceResourceId, currentApp.ClientId, currentApp.ReturnUri);
                    }

                    // Check the result of the authentication operation
                    if (authenticationResult.Status != AuthenticationStatus.Success)
                    {
                        // Something went wrong, probably the user cancelled the sign in process
                        return null;
                    }

                    // Store relevant info about user and resource
                    _loggedInUser = authenticationResult.UserInfo.UniqueId;
                    _userAccount = authenticationResult.UserInfo.DisplayableId;
                    _serviceResourceId = serviceResourceId;

                    // Store relevant info about the token
                    _accessToken = cacheItem.AccessToken;
                    _accessTokenExpiresOn = cacheItem.ExpiresOn;
                    _refreshToken = cacheItem.RefreshToken;

                    // The access token is part of the data source. 
                    // We should update the data source whenever the token changes
                    //Update the data source
                    AccessTokenChanged(null, EventArgs.Empty);

                    return _accessToken;
                }
                // The following is a list of all exceptions you should consider handling in your app.
                // In the case of this sample, the exceptions are handled by returning null upstream. 
                catch (MissingConfigurationValueException mcve)
                {
                    MessageDialogHelper.DisplayException(mcve);

                    // Connected services not added correctly, or permissions not set correctly.
                    _authenticationContext.TokenCache.Clear();
                    return null;
                }
                catch (AuthenticationFailedException afe)
                {
                    MessageDialogHelper.DisplayException(afe);

                    // Failed to authenticate the user
                    _authenticationContext.TokenCache.Clear();
                    return null;

                }
                catch (ArgumentException ae)
                {
                    MessageDialogHelper.DisplayException(ae as Exception);

                    // Argument exception
                    _authenticationContext.TokenCache.Clear();
                    return null;
                }
            }
        }

        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        public static async Task SignOutAsync()
        {
            if (string.IsNullOrEmpty(_loggedInUser))
            {
                return;
            }

            await _authenticationContext.LogoutAsync(_loggedInUser);

            //Clear the cache
            _authenticationContext.TokenCache.Clear();

            // Destroy or initialize objects
            _accessToken = null;
            _accessTokenExpiresOn = DateTimeOffset.MinValue.AddSeconds(10);
            _refreshToken = null;
            _loggedInUser = null;
            _serviceResourceId = null;
            _userAccount = null;
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