/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using Office365RESTExplorerforSites.Common;
using Office365RESTExplorerforSites.Helpers;
using System;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;

namespace Office365RESTExplorerforSites
{
    /// <summary>
    /// A settings flyout that shows the configured username and site
    /// </summary>
    public sealed partial class AccountFlyout : SettingsFlyout
    {
        private ObservableDictionary defaultViewModel = new ObservableDictionary();
        public AccountFlyout()
        {
            this.InitializeComponent();

            // Add the local settings to the view model
            this.DefaultViewModel["ServiceResourceId"] = AuthenticationHelper.ServiceResourceId;
            this.DefaultViewModel["UserAccount"] = AuthenticationHelper.UserAccount;
            
            if(!String.IsNullOrEmpty(AuthenticationHelper.ServiceResourceId))
            {
                this.DefaultViewModel["SignOutVisible"] = Visibility.Visible;
            }
            else
            {
                this.DefaultViewModel["SignOutVisible"] = Visibility.Collapsed;
            }
        }

        /// <summary>
        /// This can be changed to a strongly typed view model.
        /// </summary>
        public ObservableDictionary DefaultViewModel
        {
            get { return this.defaultViewModel; }
        }

        private async void SignOut_Click(object sender, RoutedEventArgs e)
        {
            await AuthenticationHelper.SignOutAsync();

            // Add the local settings to the view model
            this.DefaultViewModel["ServiceResourceId"] = null;
            this.DefaultViewModel["UserAccount"] = null;
            this.DefaultViewModel["SignOutVisible"] = Visibility.Collapsed;

            (Window.Current.Content as Frame).Navigate(typeof(ConfigureAppPage));
        }

    }
}
