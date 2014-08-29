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

            Binding serviceResourceIdBinding = new Binding();
            serviceResourceIdBinding.Mode = BindingMode.OneWay;
            serviceResourceIdBinding.Source = ApplicationData.Current.LocalSettings.Values["ServiceResourceId"];
            this.txtSite.SetBinding(TextBlock.TextProperty, serviceResourceIdBinding);
            Binding userAccountBinding = new Binding();
            userAccountBinding.Mode = BindingMode.OneWay;
            userAccountBinding.Source = ApplicationData.Current.LocalSettings.Values["ServiceResourceId"];
            this.txtUser.SetBinding(TextBlock.TextProperty, userAccountBinding);
        }
    }
}
