/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using Microsoft.Office365.OAuth;
using Office365RESTExplorerforSites.Common;
using Office365RESTExplorerforSites.Data;
using Office365RESTExplorerforSites.Helpers;
using System;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Navigation;

namespace Office365RESTExplorerforSites
{
    /// <summary>
    /// A page that shows the user a textbox for the SharePoint site to use
    /// This page should only run once when the user installs the app
    /// </summary>
    public sealed partial class ConfigureAppPage : Page
    {
        private NavigationHelper navigationHelper;

        /// <summary>
        /// NavigationHelper is used on each page to aid in navigation and 
        /// process lifetime management
        /// </summary>
        public NavigationHelper NavigationHelper
        {
            get { return this.navigationHelper; }
        }

        public ConfigureAppPage()
        {
            this.InitializeComponent();
            this.navigationHelper = new NavigationHelper(this);
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            Uri spSiteUri;

            try
            {
                //Validate that the input is at least a well-formed URI
                spSiteUri = new Uri(spSite.Text);

                await AuthenticationHelper.EnsureAccessTokenAvailableAsync(spSiteUri.AbsoluteUri);

                this.Frame.Navigate(typeof(ItemsPage));
            }
            catch (FormatException fe)
            {
                // Tell the user that the authentication failed
                MessageDialogHelper.DisplayException(fe);
                return;
            }
            catch (AuthenticationFailedException afe)
            {
                // Tell the user that the authentication failed
                MessageDialogHelper.DisplayException(afe);
                return;
            }
        }

        #region NavigationHelper registration

        /// The methods provided in this section are simply used to allow
        /// NavigationHelper to respond to the page's navigation methods.
        /// 
        /// Page specific logic should be placed in event handlers for the  
        /// <see cref="GridCS.Common.NavigationHelper.LoadState"/>
        /// and <see cref="GridCS.Common.NavigationHelper.SaveState"/>.
        /// The navigation parameter is available in the LoadState method 
        /// in addition to page state preserved during an earlier session.

        protected async override void OnNavigatedTo(NavigationEventArgs e)
        {
            navigationHelper.OnNavigatedTo(e);

            // If I'm starting the app or returning to this page, it means that 
            // I want to sign in again
            await AuthenticationHelper.SignOutAsync();
            DataSource.Clear();
        }

        protected override void OnNavigatedFrom(NavigationEventArgs e)
        {
            navigationHelper.OnNavigatedFrom(e);
        }

        #endregion
    }
}
