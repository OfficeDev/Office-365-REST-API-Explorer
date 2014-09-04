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

namespace Office365RESTExplorerforSites
{
    /// <summary>
    /// A settings flyout that shows the configured username and site
    /// </summary>
    public sealed partial class SettingsFlyout1 : SettingsFlyout
    {
        public SettingsFlyout1()
        {
            this.InitializeComponent();

            //Create the data bindings
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
