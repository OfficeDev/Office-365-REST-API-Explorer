using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Threading.Tasks;
using Windows.Storage;
using Office365RESTExplorerforSites.Common;
using Microsoft.Office365.OAuth;
using Windows.UI.Popups;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace Office365RESTExplorerforSites
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class StartPage : Page
    {
        public StartPage()
        {
            this.InitializeComponent();

            Binding siteBinding = new Binding();
            siteBinding.Mode = BindingMode.TwoWay;
            siteBinding.Source = ApplicationData.Current.LocalSettings.Values["ServiceResourceId"];
            siteBinding.FallbackValue = "Office 365 site e.g. https://contoso.sharepoint.com";
            this.spSite.SetBinding(TextBox.TextProperty, siteBinding);
        }



        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            Uri spSiteUri;
            DiscoveryContext _discoveryContext;
            AuthenticationResult authResult;
            MessageDialog errorDialog = null;

            try
            {
                //Validate that the input is at least a well-formed URI
                spSiteUri = new Uri(spSite.Text);

                _discoveryContext = await DiscoveryContext.CreateAsync();

                var dcr = await _discoveryContext.DiscoverResourceAsync(spSiteUri.AbsoluteUri);

                authResult = await _discoveryContext.AuthenticationContext.AcquireTokenSilentAsync(
                                                                                spSiteUri.AbsoluteUri,
                                                                                _discoveryContext.AppIdentity.ClientId,
                                                                                new UserIdentifier(dcr.UserId, UserIdentifierType.UniqueId)
                                                                                );


                if (authResult.Status != AuthenticationStatus.Success)
                {
                    throw new AuthenticationFailedException(authResult.Error, authResult.ErrorDescription);
                }

                // Store the relevant data in local settings.
                ApplicationData.Current.LocalSettings.Values["ServiceResourceId"] = spSiteUri.AbsoluteUri;
                ApplicationData.Current.LocalSettings.Values["UserId"] = dcr.UserId;
                ApplicationData.Current.LocalSettings.Values["UserAccount"] = authResult.UserInfo.DisplayableId;
                ApplicationData.Current.LocalSettings.Values["AccessToken"] = authResult.AccessToken;
                ApplicationData.Current.LocalSettings.Values["RefreshToken"] = authResult.RefreshToken;
                ApplicationData.Current.LocalSettings.Values["AccessTokenExpiresOn"] = authResult.ExpiresOn;

                this.Frame.Navigate(typeof(ItemsPage));
            }
            catch (FormatException)
            {
                // Tell the user to correct the site URL
                errorDialog = new MessageDialog("It looks like the Office 365 site is not a valid URL.", "Invalid Office 365 site");
            }
            catch (AuthenticationFailedException)
            {
                // Tell the user that the authentication failed
                errorDialog = new MessageDialog("We couldn't sign you in to your Office 356 site.", "Authentication failed");
            }

            if (errorDialog != null)
                await errorDialog.ShowAsync();

        }
    }
}
