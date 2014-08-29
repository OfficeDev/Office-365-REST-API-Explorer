using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office365.OAuth;
using Microsoft.Office365.SharePoint;
using Windows.Storage;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

static class Office365Helper
    {
        private static DiscoveryContext _discoveryContext;

        public static async Task<String[]> AcquireAccessTokenAsync(string ServiceResourceId)
        {
            UserIdentifier _userIdObj;
            AuthenticationResult authResult = null;
            bool tokenInCache = false;
            TokenCacheItem tci = null;

            //Validate that the ServiceResourceId parameter is at least a well-formed URI
            if (!Uri.IsWellFormedUriString(ServiceResourceId, UriKind.Absolute))
            {
                throw new FormatException("The ServiceResourceId parameter is not a well-formed URI string.");
            }

            if (_discoveryContext == null)
            {
                _discoveryContext = await DiscoveryContext.CreateAsync();
            }

            if (_discoveryContext.LastLoggedInUser == null)
            {
                // This is the first time the user starts the app, full sign in
                ResourceDiscoveryResult dcr = await _discoveryContext.DiscoverResourceAsync(ServiceResourceId);
                _userIdObj = new UserIdentifier(dcr.UserId, UserIdentifierType.UniqueId);
                authResult = await _discoveryContext.AuthenticationContext.AcquireTokenSilentAsync(ServiceResourceId, _discoveryContext.AppIdentity.ClientId, _userIdObj);
                if (authResult.Status == AuthenticationStatus.Success)
                {
                    // We have a new token, the user explicitly authenticated
                    return new string[] { authResult.AccessToken, authResult.UserInfo.UniqueId, authResult.UserInfo.DisplayableId};
                }
                else
                {
                    throw new AuthenticationFailedException(authResult.Error, authResult.StatusCode.ToString(), authResult.ErrorDescription);
                }
            }

            try
            {
                tci = await GetTokenFromCacheAsync(ServiceResourceId, _discoveryContext.LastLoggedInUser);
                tokenInCache = true;
            }
            catch (KeyNotFoundException)
            {
                //There are no tokens for this user/resource pair in the cache. Authenticate the user.
                // You cannot call an awaitable function from a catch clause. Set up a flag instead.
                // tokenInCache is our flag, setting it to false to flag that the token was not found.
                tokenInCache = false;
            }
            if (!tokenInCache)
            {
                ResourceDiscoveryResult dcr = await _discoveryContext.DiscoverResourceAsync(ServiceResourceId);
                _userIdObj = new UserIdentifier(dcr.UserId, UserIdentifierType.UniqueId);
                authResult = await _discoveryContext.AuthenticationContext.AcquireTokenSilentAsync(ServiceResourceId, _discoveryContext.AppIdentity.ClientId, _userIdObj);
                if (authResult.Status == AuthenticationStatus.Success)
                {
                    // We have a new token from the refresh flow
                    return new string[] { authResult.AccessToken, authResult.UserInfo.UniqueId, authResult.UserInfo.DisplayableId };
                }
                else
                {
                    throw new AuthenticationFailedException(authResult.Error, authResult.StatusCode.ToString(), authResult.ErrorDescription);
                }
            }

            // We have a matching token in the cache, has it expired?
            if (DateTimeOffset.Compare(tci.ExpiresOn, DateTimeOffset.Now) <= 0)
            {
                // The token has expired go to the refresh flow
                authResult = await _discoveryContext.AuthenticationContext.AcquireTokenByRefreshTokenAsync(tci.RefreshToken, tci.ClientId, tci.Resource);
                if (authResult.Status == AuthenticationStatus.Success)
                {
                    // We have a new token from the refresh flow
                    return new string[] { authResult.AccessToken, tci.UniqueId, tci.DisplayableId };
                }
                else
                {
                    // We couldn't refresh the token. It could have been revoked or the refresh token expired (unlikely but possible).
                    ResourceDiscoveryResult dcr = await _discoveryContext.DiscoverResourceAsync(ServiceResourceId);
                    _userIdObj = new UserIdentifier(dcr.UserId, UserIdentifierType.UniqueId);
                    authResult = await _discoveryContext.AuthenticationContext.AcquireTokenSilentAsync(ServiceResourceId, _discoveryContext.AppIdentity.ClientId, _userIdObj);
                    if (authResult.Status == AuthenticationStatus.Success)
                    {
                        // We have a new token, the user explicitly authenticated
                        return new string[] { authResult.AccessToken, authResult.UserInfo.UniqueId, authResult.UserInfo.DisplayableId };
                    }
                    else
                    {
                        throw new AuthenticationFailedException(authResult.Error, authResult.StatusCode.ToString(), authResult.ErrorDescription);
                    }
                }
            }
            
            // Most of the cases should be handled here. The user has a valid token in cache.
            _userIdObj = new UserIdentifier(_discoveryContext.LastLoggedInUser, UserIdentifierType.UniqueId);
            
            return new string[] { tci.AccessToken, tci.UniqueId, tci.DisplayableId };
        }

        public static async Task LogoutAsync(string UserIdentifier)
        {
            if (_discoveryContext == null)
            {
                _discoveryContext = await DiscoveryContext.CreateAsync();
            }
            await _discoveryContext.LogoutAsync(UserIdentifier);
            _discoveryContext.AuthenticationContext.TokenCache.Clear();
        }

        public static async Task<TokenCacheItem> GetTokenFromCacheAsync(string ServiceResourceId, string UserIdentifier)
        {
            if (_discoveryContext == null)
            {
                _discoveryContext = await DiscoveryContext.CreateAsync();
            }

            IEnumerable<TokenCacheItem> tci = _discoveryContext.AuthenticationContext.TokenCache.ReadItems();

            foreach (TokenCacheItem item in tci)
            {
                bool resourceMatches = ServiceResourceId.Equals(new Uri(item.Resource));
                bool userIdMatches = String.Compare(item.UniqueId, UserIdentifier, StringComparison.CurrentCultureIgnoreCase) == 0;

                if (resourceMatches && userIdMatches)
                    return item;
            }
            throw new KeyNotFoundException("The token was not found in the cache.");
        }

        public static async Task ClearTokenCacheAsync()
        {
            if (_discoveryContext == null)
            {
                _discoveryContext = await DiscoveryContext.CreateAsync();
            }

            _discoveryContext.AuthenticationContext.TokenCache.Clear();
        }
    }

