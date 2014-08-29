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
            try
            {
                Uri spSiteUri = new Uri(spSite.Text);
            }
            catch (FormatException fe)
            {

            }


            DiscoveryContext _discoveryContext = await DiscoveryContext.CreateAsync();

            if (ApplicationData.Current.LocalSettings.Values["UserId"] != null)
            {
                await _discoveryContext.LogoutAsync(ApplicationData.Current.LocalSettings.Values["UserId"].ToString());
                _discoveryContext = await DiscoveryContext.CreateAsync();
            }

            var dcr = await _discoveryContext.DiscoverResourceAsync(spSite.Text);

            AuthenticationResult authResult = await _discoveryContext.AuthenticationContext.AcquireTokenSilentAsync(spSite.Text, _discoveryContext.AppIdentity.ClientId, new Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier(dcr.UserId, Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifierType.UniqueId));

            
            ApplicationData.Current.LocalSettings.Values["SharePointSiteUrl"] = 
            ApplicationData.Current.LocalSettings.Values["ServiceResourceId"] = spSite.Text;
            ApplicationData.Current.LocalSettings.Values["UserId"] = dcr.UserId;
            ApplicationData.Current.LocalSettings.Values["UserAccount"] = authResult.UserInfo.DisplayableId;
            ApplicationData.Current.LocalSettings.Values["AccessToken"] = authResult.AccessToken;

            this.Frame.Navigate(typeof(ItemsPage));
        }
    }
}
