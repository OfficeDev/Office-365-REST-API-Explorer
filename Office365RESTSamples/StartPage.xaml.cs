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

using Microsoft.Office365.OAuth;
using System.Threading.Tasks;
using Windows.Storage;

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
            if (ApplicationData.Current.LocalSettings.Values["ServiceResourceId"] != null)
                spSite.Text = ApplicationData.Current.LocalSettings.Values["ServiceResourceId"].ToString();
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            Uri uriSite = new Uri(spSite.Text);
            await Office365Helper.SignIn(uriSite);

            this.Frame.Navigate(typeof(ItemsPage));
        }
    }
}
