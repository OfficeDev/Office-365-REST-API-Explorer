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

using Windows.Storage;
using Office365RESTExplorerforSites.Common;
using System.Threading.Tasks;
using Office365RESTExplorerforSites.Helpers;

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
