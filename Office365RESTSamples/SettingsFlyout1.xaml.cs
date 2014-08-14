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

using Windows.Storage;
using Microsoft.Office365.OAuth;

// The Settings Flyout item template is documented at http://go.microsoft.com/fwlink/?LinkId=273769

namespace Office365RESTExplorerforSites
{
    public sealed partial class SettingsFlyout1 : SettingsFlyout
    {
        public SettingsFlyout1()
        {
            this.InitializeComponent();

            if (ApplicationData.Current.LocalSettings.Values["UserAccount"] != null)
            {
                stkSignedIn.Visibility = Windows.UI.Xaml.Visibility.Visible;
                stkSignedOut.Visibility = Windows.UI.Xaml.Visibility.Collapsed;
                txtSite.Text = ApplicationData.Current.LocalSettings.Values["ServiceResourceId"].ToString();
                txtUser.Text = ApplicationData.Current.LocalSettings.Values["UserAccount"].ToString();
            }
            else
            {
                stkSignedIn.Visibility = Windows.UI.Xaml.Visibility.Collapsed;
                stkSignedOut.Visibility = Windows.UI.Xaml.Visibility.Visible;
                txtNewSite.Text = ApplicationData.Current.LocalSettings.Values["ServiceResourceId"].ToString();
            }
        }

        private async void Logout_Click(object sender, RoutedEventArgs e)
        {
            await Office365Helper.Logout();

            txtNewSite.Text = ApplicationData.Current.LocalSettings.Values["ServiceResourceId"].ToString();
            stkSignedIn.Visibility = Windows.UI.Xaml.Visibility.Collapsed;
            stkSignedOut.Visibility = Windows.UI.Xaml.Visibility.Visible;
        }

        private async void SignIn_Click(object sender, RoutedEventArgs e)
        {
            
            Uri spSite = new Uri(txtNewSite.Text);
            await Office365Helper.SignIn(spSite);
            stkSignedIn.Visibility = Windows.UI.Xaml.Visibility.Visible;
            stkSignedOut.Visibility = Windows.UI.Xaml.Visibility.Collapsed;
        }
    }
}
