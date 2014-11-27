// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using Microsoft.Office365.OAuth;
using System;
using System.Threading.Tasks;
using Windows.Storage;

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

        // Private property used to store the access token
        // We need to detect if the access token changed.
        private static string _accessToken
        {
            get
            {
                if (ApplicationData.Current.LocalSettings.Values.ContainsKey("AccessToken")
                    && ApplicationData.Current.LocalSettings.Values["AccessToken"] != null)
                {
                    return ApplicationData.Current.LocalSettings.Values["AccessToken"].ToString();
                }
                else
                {
                    return null;
                }
            }
            set
            {
                ApplicationData.Current.LocalSettings.Values["AccessToken"] = value;
            }
        }

        private static AuthenticationContext _authenticationContext;

        // Property for storing the logged-in user so that we can display user properties later.
        //This value is populated when the user connects to the service and made null when the user signs out.
        private static string _loggedInUser
        {
            get
            {
                if (ApplicationData.Current.LocalSettings.Values.ContainsKey("LoggedInUser") 
                    && ApplicationData.Current.LocalSettings.Values["LoggedInUser"] != null)
                {
                    return ApplicationData.Current.LocalSettings.Values["LoggedInUser"].ToString();
                }
                else
                {
                    return null;
                }
            }
            set
            {
                ApplicationData.Current.LocalSettings.Values["LoggedInUser"] = value;
            }
        }

        //Property for storing and returning the authority used by the last authentication.
        //This value is populated when the user connects to the service and made null when the user signs out.
        private static string _lastAuthority
        {
            get
            {
                if (ApplicationData.Current.LocalSettings.Values.ContainsKey("LastAuthority") 
                    && ApplicationData.Current.LocalSettings.Values["LastAuthority"] != null)
                {
                    return ApplicationData.Current.LocalSettings.Values["LastAuthority"].ToString();
                }
                else
                {
                    return null;
                }

            }
            set
            {
                ApplicationData.Current.LocalSettings.Values["LastAuthority"] = value;
            }
        }

        // We need an event to notify other classes that the
        // access token has changed and they need to handle this.
        // For example, they need to update the data source
        internal static event EventHandler AccessTokenChanged = delegate { };

        /// <summary> 
        /// Gets the service resource id
        /// </summary> 
        internal static string ServiceResourceId
        {
            get
            {
                if (ApplicationData.Current.LocalSettings.Values.ContainsKey("ServiceResourceId")
                    && ApplicationData.Current.LocalSettings.Values["ServiceResourceId"] != null)
                {
                    return ApplicationData.Current.LocalSettings.Values["ServiceResourceId"].ToString();
                }
                else
                {
                    return null;
                }
            }
            set
            {
                ApplicationData.Current.LocalSettings.Values["ServiceResourceId"] = value;
            }
        }

        /// <summary> 
        /// Gets the account of the user in the form user@domain.tld
        /// </summary> 
        internal static string UserAccount
        {
            get
            {
                if (ApplicationData.Current.LocalSettings.Values.ContainsKey("UserAccount")
                    && ApplicationData.Current.LocalSettings.Values["UserAccount"] != null)
                {
                    return ApplicationData.Current.LocalSettings.Values["UserAccount"].ToString();
                }
                else
                {
                    return null;
                }
            }
            set
            {
                ApplicationData.Current.LocalSettings.Values["UserAccount"] = value;
            }
        }

        /// <summary>
        /// Checks that an access token is available.
        /// This method requires that the ServiceResourceId has been set previously.
        /// </summary>
        /// <returns>The access token.</returns>
        internal static async Task<string> EnsureAccessTokenAvailableAsync()
        {
            if(!String.IsNullOrEmpty(ServiceResourceId))
            {
                return await EnsureAccessTokenAvailableAsync(ServiceResourceId);
            }
            else
            {
                MissingConfigurationValueException mcve = 
                    new MissingConfigurationValueException(
                        "To use this method you have to call EnsureAccessTokenCreatedAsync(string serviceResourceId) at least once."
                        );
                MessageDialogHelper.DisplayException(mcve);
                return null;
            }
        }

        /// <summary>
        /// Checks that an access token is available.
        /// </summary>
        /// <returns>The access token.</returns>
        internal static async Task<string> EnsureAccessTokenAvailableAsync(string serviceResourceId)
        {
            try
            {
                // First, look for the authority used during the last authentication.
                // If that value is not populated, use _commonAuthority.
                string authority = null;
                if (String.IsNullOrEmpty(_lastAuthority))
                {
                    authority = _commonAuthority;
                }
                else
                {
                    authority = _lastAuthority;
                }

                // Create an AuthenticationContext using this authority.
                _authenticationContext = new AuthenticationContext(authority, true);

                //Get the current app object, which exposes the ClientId and ReturnUri properties
                // that we need in the following call to AcquireTokenAsync
                App currentApp = (App)App.Current;
                    
                AuthenticationResult authenticationResult;

                // An attempt is first made to acquire the token silently. 
                // If that fails, then we try to acquire the token by prompting the user.
                authenticationResult = await _authenticationContext.AcquireTokenSilentAsync(serviceResourceId, currentApp.ClientId);

                if (authenticationResult.Status != AuthenticationStatus.Success)
                {
                    // Try to authenticate by prompting the user
                    authenticationResult = await _authenticationContext.AcquireTokenAsync(serviceResourceId, currentApp.ClientId, currentApp.ReturnUri);
                    
                    // Check the result of the authentication operation
                    if (authenticationResult.Status != AuthenticationStatus.Success)
                    {
                        // Something went wrong, probably the user cancelled the sign in process
                        return null;
                    }
                }

                // Store relevant info about user and resource
                _loggedInUser = authenticationResult.UserInfo.UniqueId;
                // The new last authority is in the form https://login.windows.net/{TenantId}
                _lastAuthority = App.Current.Resources["ida:AuthorizationUri"].ToString() + "/" + authenticationResult.TenantId;
                UserAccount = authenticationResult.UserInfo.DisplayableId;
                ServiceResourceId = serviceResourceId;

                // If the acccess token has changed
                if (!String.Equals(_accessToken, authenticationResult.AccessToken))
                {
                    // Raise an event to let other components know that the token has changed, 
                    // so they can react accordingly (for example, updating the data source)
                    AccessTokenChanged(null, EventArgs.Empty);
                    // and store the new acces token
                    _accessToken = authenticationResult.AccessToken;
                }

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

        /// <summary>
        /// Signs the user out of the service.
        /// </summary>
        internal static async Task SignOutAsync()
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
            _loggedInUser = null;
            _lastAuthority = null;
            ServiceResourceId = null;
            UserAccount = null;
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