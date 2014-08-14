using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office365.OAuth;
using Microsoft.Office365.SharePoint;
using Windows.Storage;
using Microsoft.IdentityModel.Clients.ActiveDirectory;

namespace Office365RESTExplorerforSites
{
    static class Office365Helper
    {
        public static DiscoveryContext _discoveryContext;
        public static async Task SignIn(Uri ServiceResourceId)
        {
            Uri ServiceEndpointUri = new Uri(ServiceResourceId.AbsoluteUri + "_api/");
            bool tokensFoundinCache = false;
            AuthenticationResult authResult;
            TokenCacheItem tci = null;

            if (_discoveryContext == null)
            {
                _discoveryContext = await DiscoveryContext.CreateAsync();
            }

            try
            {
                tci = await Office365Helper.GetTokenFromCache();
                
                if(DateTimeOffset.Compare(tci.ExpiresOn, DateTimeOffset.Now) <= 0) //If the token has expired.
                {
                    //Get another one with the refreshToken
                    authResult = await _discoveryContext.AuthenticationContext.AcquireTokenByRefreshTokenAsync(tci.RefreshToken, tci.ClientId, tci.Resource);
                    tci = await Office365Helper.GetTokenFromCache();
                }

                tokensFoundinCache = true;
            }
            catch (KeyNotFoundException)
            {
                //TODO: We need tokens, set this flag to false
                tokensFoundinCache = false;
            }

            if (!tokensFoundinCache) //TODO: we might need to validate if the tokens are invalid
            {
                ResourceDiscoveryResult dcr = await _discoveryContext.DiscoverResourceAsync(ServiceResourceId.AbsoluteUri);
                authResult = await _discoveryContext.AuthenticationContext.AcquireTokenSilentAsync(ServiceResourceId.AbsoluteUri, _discoveryContext.AppIdentity.ClientId, new Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier(dcr.UserId, Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifierType.UniqueId));
                tci = await Office365Helper.GetTokenFromCache();
            }

            ApplicationData.Current.LocalSettings.Values["UserAccount"] = tci.DisplayableId;
            ApplicationData.Current.LocalSettings.Values["ServiceResourceId"] = tci.Resource;
        }

        public static async Task Logout()
        {
            if (ApplicationData.Current.LocalSettings.Values["UserAccount"] == null)
            {
                return;
            }

            if (_discoveryContext == null)
            {
                _discoveryContext = await DiscoveryContext.CreateAsync();
            }
            await _discoveryContext.LogoutAsync(ApplicationData.Current.LocalSettings.Values["UserAccount"].ToString());
            //_discoveryContext.AuthenticationContext.TokenCache.Clear();
            //ApplicationData.Current.LocalSettings.Values.Remove("UserAccount");
        }

        public static async Task<TokenCacheItem> GetTokenFromCache()
        {
            if (_discoveryContext == null)
            {
                _discoveryContext = await DiscoveryContext.CreateAsync();
            }

            IEnumerable<TokenCacheItem> tci = _discoveryContext.AuthenticationContext.TokenCache.ReadItems();

            foreach (TokenCacheItem item in tci)
            {
                if (item.Resource == ApplicationData.Current.LocalSettings.Values["ServiceResourceId"].ToString()) //item.DisplayableId == ApplicationData.Current.LocalSettings.Values["UserAccount"].ToString() &&
                    return item;
            }
            throw new KeyNotFoundException("The token was not found in the cache.");
        }
    }
}
