using Office365RESTExplorerforSites.Common;
using Office365RESTExplorerforSites.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Windows.Input;
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
using Windows.UI.Popups;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Threading.Tasks;
using Windows.Storage;
using Microsoft.Office365.OAuth;

// The Items Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234233

namespace Office365RESTExplorerforSites
{
    /// <summary>
    /// A page that displays a collection of item previews.  In the Split App this page
    /// is used to display and select one of the available groups.
    /// </summary>
    public sealed partial class ItemsPage : Page
    {
        private NavigationHelper navigationHelper;
        private ObservableDictionary defaultViewModel = new ObservableDictionary();
        private const int MinimumWidthForSupportingTwoPanes = 768;

        /// <summary>
        /// NavigationHelper is used on each page to aid in navigation and 
        /// process lifetime management
        /// </summary>
        public NavigationHelper NavigationHelper
        {
            get { return this.navigationHelper; }
        }

        /// <summary>
        /// This can be changed to a strongly typed view model.
        /// </summary>
        public ObservableDictionary DefaultViewModel
        {
            get { return this.defaultViewModel; }
        }

        public ItemsPage()
        {
            this.InitializeComponent();
            this.navigationHelper = new NavigationHelper(this);
            this.navigationHelper.LoadState += navigationHelper_LoadState;
        }

        
        /// <summary>
        /// Populates the page with content passed during navigation.  Any saved state is also
        /// provided when recreating a page from a prior session.
        /// </summary>
        /// <param name="sender">
        /// The source of the event; typically <see cref="NavigationHelper"/>
        /// </param>
        /// <param name="e">Event data that provides both the navigation parameter passed to
        /// <see cref="Frame.Navigate(Type, Object)"/> when this page was initially requested and
        /// a dictionary of state preserved by this page during an earlier
        /// session.  The state will be null the first time a page is visited.</param>
        private async void navigationHelper_LoadState(object sender, LoadStateEventArgs e)
        {
            var sampleDataGroups = await DataSource.GetGroupsAsync();
            this.DefaultViewModel["Items"] = sampleDataGroups;
        }

        /// <summary>
        /// Invoked when an item is clicked.
        /// </summary>
        /// <param name="sender">The GridView (or ListView when the application is snapped)
        /// displaying the item clicked.</param>
        /// <param name="e">Event data that describes the item clicked.</param>
        void ItemView_ItemClick(object sender, ItemClickEventArgs e)
        {
            // Navigate to the appropriate destination page, configuring the new page
            // by passing required information as a navigation parameter
            var groupId = ((DataGroup)e.ClickedItem).UniqueId;
            this.Frame.Navigate(typeof(SplitPage), groupId);
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

        protected override async void OnNavigatedTo(NavigationEventArgs e)
        {
            navigationHelper.OnNavigatedTo(e);

            // If the access token has expired, renew it and update the data source
            if (DateTimeOffset.Compare((DateTimeOffset)ApplicationData.Current.LocalSettings.Values["AccessTokenExpiresOn"], DateTimeOffset.Now) <= 0)
            {
                DiscoveryContext _discoveryContext;
                AuthenticationResult authResult;
                MessageDialog errorDialog = null;
                

                try
                {
                    Uri spSiteUri = new Uri(ApplicationData.Current.LocalSettings.Values["ServiceResourceId"].ToString());
                    _discoveryContext = await DiscoveryContext.CreateAsync();

                    var dcr = await _discoveryContext.DiscoverResourceAsync(spSiteUri.AbsoluteUri);

                    authResult = await _discoveryContext.AuthenticationContext.AcquireTokenSilentAsync(
                                                                                    spSiteUri.AbsoluteUri,
                                                                                    _discoveryContext.AppIdentity.ClientId,
                                                                                    new UserIdentifier(dcr.UserId, UserIdentifierType.UniqueId)
                                                                                    );


                    if (authResult.Status != AuthenticationStatus.Success)
                    {
                        throw new AuthenticationFailedException(authResult.Error, authResult.ErrorDescription);
                    }

                    // Store the relevant data in local settings.
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

                //Update the data source
                var sampleDataGroups = await DataSource.GetGroupsAsync();
                this.DefaultViewModel["Items"] = sampleDataGroups;
            }
        }

        protected override void OnNavigatedFrom(NavigationEventArgs e)
        {
            navigationHelper.OnNavigatedFrom(e);
        }

        #endregion

        private void pageRoot_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            // If the window resizes go to the visual state that works for the width
            //VisualStateManager.GoToState(this, DetermineVisualState(), false);
            if(Window.Current.Bounds.Width < MinimumWidthForSupportingTwoPanes)
            {
                secondaryColumn.Width = new GridLength(0);
                Description.Visibility = Windows.UI.Xaml.Visibility.Collapsed;
            }
            else
            {
                secondaryColumn.Width = new GridLength(1, GridUnitType.Star);
                Description.Visibility = Windows.UI.Xaml.Visibility.Visible;
            }
        }

        /// <summary>
        /// Invoked to determine the name of the visual state that corresponds to an application
        /// view state.
        /// </summary>
        /// <returns>The name of the desired visual state.  This is the same as the name of the
        /// view state except when there is a selected item in portrait and snapped views where
        /// this additional logical page is represented by adding a suffix of _Detail.</returns>
        private string DetermineVisualState()
        {
            //If the width is less than the supported minimum reurn SinglePane
            return Window.Current.Bounds.Width < MinimumWidthForSupportingTwoPanes ? "PrimaryView" : "SinglePane";
        }
    }
}
