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
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            String[] authResult = await Office365Helper.AcquireAccessTokenAsync(spSite.Text);

            ApplicationData.Current.LocalSettings.Values["ServiceResourceId"] = spSite.Text;
            ApplicationData.Current.LocalSettings.Values["AccessToken"] = authResult[0];
            ApplicationData.Current.LocalSettings.Values["UserId"] = authResult[1]; 
            ApplicationData.Current.LocalSettings.Values["UserAccount"] = authResult[2];

            this.Frame.Navigate(typeof(ItemsPage));
        }

       

    }
}
