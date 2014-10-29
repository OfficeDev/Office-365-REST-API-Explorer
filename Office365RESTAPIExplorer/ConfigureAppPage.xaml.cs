// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

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
using Windows.Storage;
using Microsoft.Office365.OAuth;
using Windows.UI.Popups;

namespace Office365RESTExplorerforSites
{
    /// <summary>
    /// A page that shows the user a textbox for the SharePoint site to use
    /// This page should only run once when the user installs the app
    /// </summary>
    public sealed partial class ConfigureAppPage : Page
    {
        public ConfigureAppPage()
        {
            this.InitializeComponent();
        }

        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            Uri spSiteUri;
            AuthenticationResult authResult;
            MessageDialog errorDialog = null;

            try
            {
                //Validate that the input is at least a well-formed URI
                spSiteUri = new Uri(spSite.Text);

                authResult = await AuthenticationHelper.EnsureAccessTokenAvailableAsync(spSiteUri.AbsoluteUri, PromptBehavior.Auto);

                if (authResult.Status != AuthenticationStatus.Success)
                {
                    throw new AuthenticationFailedException(authResult.Error, authResult.ErrorDescription);
                }

                // Store the relevant data in local settings.
                ApplicationData.Current.LocalSettings.Values["ServiceResourceId"] = spSiteUri.AbsoluteUri;
                ApplicationData.Current.LocalSettings.Values["UserId"] = authResult.UserInfo.UniqueId;
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