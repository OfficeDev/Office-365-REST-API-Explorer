using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using Windows.Data.Json;
using Windows.Storage;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Media.Imaging;

using Newtonsoft.Json;
using Windows.UI.Xaml.Data;
using System.IO;
using System.Net;
using System.Text;
using System.ComponentModel;
using System.Runtime.CompilerServices;

// The data model defined by this file serves as a representative example of a strongly-typed
// model.  The property names chosen coincide with data bindings in the standard item templates.
//
// Applications may use this model as a starting point and build on it, or discard it entirely and
// replace it with something appropriate to their needs. If using this model, you might improve app 
// responsiveness by initiating the data loading task in the code behind for App.xaml when the app 
// is first launched.

namespace Office365RESTExplorerforSites.Data
{
    public class ResponseItem
    {
        public ResponseItem(string responseUri, string status, JsonObject headers, JsonObject body)
        {
            this.ResponseUri = responseUri;
            this.Status = status;
            this.Headers = headers;
            this.Body = body;
        }

        public string ResponseUri { get; private set; }
        public JsonObject Headers { get; private set; }
        public JsonObject Body { get; private set; }
        public string Status { get; private set; }
    }
    public class RequestItem : INotifyPropertyChanged
    {
        public RequestItem(string apiUrl, string method, JsonObject headers, JsonObject body)
        {
            this.ApiUrl = apiUrl;

            // Validate that the method is either "GET" or "POST"
            if (String.Compare(method, "GET", StringComparison.CurrentCultureIgnoreCase) != 0 && String.Compare(method, "POST", StringComparison.CurrentCultureIgnoreCase) != 0)
                throw new ArgumentOutOfRangeException("The HTTP method can only be GET or POST.");
            else
                this.Method = method;

            this.Headers = headers;
            this.Body = body;
        }

        public string ApiUrl { get; set; }
        public JsonObject Headers { get; set; }
        public JsonObject Body { get; set; }
        public string Method { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        // This method is called by the Set accessor of each property. 
        // The CallerMemberName attribute that is applied to the optional propertyName 
        // parameter causes the property name of the caller to be substituted as an argument. 
        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
    /// <summary>
    /// Generic item data model.
    /// </summary>
    public class DataItem : INotifyPropertyChanged
    {
        private ResponseItem response;
        public DataItem(String uniqueId, String title, String subtitle, String imagePath)
        {
            this.UniqueId = uniqueId;
            this.Title = title;
            this.Subtitle = subtitle;
            this.ImagePath = imagePath;
        }

        public string UniqueId { get; private set; }
        public string Title { get; private set; }
        public string Subtitle { get; private set; }
        public string ImagePath { get; private set; }
        public string ApiUrl { get; private set; }
        public RequestItem Request { get; set; }
        public ResponseItem Response {
            get
            {
                return response;
            }
            set
            {
                response = value;
                // Notify the UI that the response property has changed.
                PropertyChangedEventHandler handler = PropertyChanged;
                PropertyChangedEventArgs e = new PropertyChangedEventArgs("Response");
                if (handler != null)
                {
                    handler(this, e);
                }
            } 
        }

        public override string ToString()
        {
            return this.Title;
        }

        public event PropertyChangedEventHandler PropertyChanged;

        // This method is called by the Set accessor of each property. 
        // The CallerMemberName attribute that is applied to the optional propertyName 
        // parameter causes the property name of the caller to be substituted as an argument. 
        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }

    /// <summary>
    /// Generic group data model.
    /// </summary>
    public class DataGroup
    {
        public DataGroup(String uniqueId, String title, String subtitle, String imagePath, String moreInfoText, String moreInfoUri)
        {
            this.UniqueId = uniqueId;
            this.Title = title;
            this.Subtitle = subtitle;
            this.ImagePath = imagePath;
            this.MoreInfoText = moreInfoText;
            this.MoreInfoUri = moreInfoUri;
            this.Items = new ObservableCollection<DataItem>();
        }

        public string UniqueId { get; private set; }
        public string Title { get; private set; }
        public string Subtitle { get; private set; }
        public string ImagePath { get; private set; }
        public string MoreInfoUri { get; private set; }
        public string MoreInfoText { get; private set; }
        public ObservableCollection<DataItem> Items { get; private set; }

        public override string ToString()
        {
            return this.Title;
        }
    }

    /// <summary>
    /// Creates a collection of groups and items with content read from a static json file.
    /// 
    /// SampleDataSource initializes with data read from a static json file included in the 
    /// project.  This provides sample data at both design-time and run-time.
    /// </summary>
    public sealed class DataSource
    {
        private static DataSource _sampleDataSource = new DataSource();

        private ObservableCollection<DataGroup> _groups = new ObservableCollection<DataGroup>();
        public ObservableCollection<DataGroup> Groups
        {
            get { return this._groups; }
        }

        public static async Task<IEnumerable<DataGroup>> GetGroupsAsync()
        {
            await _sampleDataSource.GetSampleDataAsync();

            return _sampleDataSource.Groups;
        }

        public static async Task<DataGroup> GetGroupAsync(string uniqueId)
        {
            await _sampleDataSource.GetSampleDataAsync();
            // Simple linear search is acceptable for small data sets
            var matches = _sampleDataSource.Groups.Where((group) => group.UniqueId.Equals(uniqueId));
            if (matches.Count() == 1) return matches.First();
            return null;
        }

        public static async Task<DataItem> GetItemAsync(string uniqueId)
        {
            await _sampleDataSource.GetSampleDataAsync();
            // Simple linear search is acceptable for small data sets
            var matches = _sampleDataSource.Groups.SelectMany(group => group.Items).Where((item) => item.UniqueId.Equals(uniqueId));
            if (matches.Count() == 1) return matches.First();
            return null;
        }

        public static async Task<ResponseItem> GetResponseAsync(RequestItem request)
        {
            HttpWebRequest endpointRequest;

            //Validate that the resulting URI is well-formed.
            Uri endpointUri = new Uri(new Uri(ApplicationData.Current.LocalSettings.Values["ServiceResourceId"].ToString()), request.ApiUrl);

            endpointRequest = (HttpWebRequest)HttpWebRequest.Create(endpointUri.AbsoluteUri);
            endpointRequest.Method = request.Method;

            // Add the headers to the request
            foreach (KeyValuePair<string, IJsonValue> header in request.Headers)
            {
                // Accept and contenttype are special cases that must be added using the Accept and ContentType properties
                // All other headers can be added using the Headers collection 
                switch (header.Key.ToLower())
                {
                    case "accept":
                        endpointRequest.Accept = header.Value.GetString();
                        break;
                    case "content-type":
                        endpointRequest.ContentType = header.Value.GetString();
                        break;
                    default:
                        endpointRequest.Headers[header.Key] = header.Value.GetString();
                        break;
                }
            }

            //Request body, added to the request only if method is POST
            if (request.Method == "POST")
            {
                string postData = request.Body.Stringify();
                UTF8Encoding encoding = new UTF8Encoding();
                byte[] byte1 = encoding.GetBytes(postData);
                System.IO.Stream newStream = await endpointRequest.GetRequestStreamAsync();
                newStream.Write(byte1, 0, byte1.Length);
            }

            Stream responseStream;
            WebHeaderCollection responseHeaders;
            string status;
            string responseUri;
            JsonObject headers = null;
            JsonObject body = null;

            try
            {
                // If the request is succesful we can use the endpointResponse object
                HttpWebResponse endpointResponse = (HttpWebResponse)await endpointRequest.GetResponseAsync();
                status = (int)endpointResponse.StatusCode + " - " + endpointResponse.StatusDescription;
                responseStream = endpointResponse.GetResponseStream();
                responseUri = endpointResponse.ResponseUri.AbsoluteUri;
                responseHeaders = endpointResponse.Headers;
            }
            catch (WebException we)
            {
                // If the request fails, we must use the response stream from the exception
                status = we.Message;
                responseStream = we.Response.GetResponseStream();
                responseUri = we.Response.ResponseUri.AbsoluteUri;
                responseHeaders = we.Response.Headers;
            }

            string responseString = string.Empty;
            using (StreamReader reader = new StreamReader(responseStream, Encoding.UTF8))
            {
                responseString = await reader.ReadToEndAsync();
            }

            if (!String.IsNullOrEmpty(responseString))
            {
                body = JsonObject.Parse(responseString);
            }

            headers = new JsonObject();
            for (int i = 0; i < responseHeaders.Count; i++)
            {
                string key = responseHeaders.AllKeys[i].ToString();
                headers.Add(key, JsonValue.CreateStringValue(responseHeaders[key]));
            }

            return new ResponseItem(responseUri, status, headers, body);
        }

        private async Task GetSampleDataAsync()
        {
            if (this._groups.Count != 0)
                return;

            Uri dataUri = new Uri("ms-appx:///DataModel/InitialData.json");

            StorageFile file = await StorageFile.GetFileFromApplicationUriAsync(dataUri);
            string jsonText = await FileIO.ReadTextAsync(file);
            JsonObject jsonObject = JsonObject.Parse(jsonText);
            JsonArray jsonArray = jsonObject["Groups"].GetArray();

            foreach (JsonValue groupValue in jsonArray)
            {
                JsonObject groupObject = groupValue.GetObject();
                DataGroup group = new DataGroup(groupObject["UniqueId"].GetString(),
                                                            groupObject["Title"].GetString(),
                                                            groupObject["Subtitle"].GetString(),
                                                            groupObject["ImagePath"].GetString(),
                                                            groupObject["MoreInfoText"].GetString(),
                                                            groupObject["MoreInfoUri"].GetString());

                foreach (JsonValue itemValue in groupObject["Items"].GetArray())
                {
                    JsonObject itemObject = itemValue.GetObject();
                    JsonObject requestObject = itemObject["Request"].GetObject();

                    //Add the Authorization header with the access token.
                    JsonObject jsonHeaders = requestObject["Headers"].GetObject();
                    jsonHeaders["Authorization"] = JsonValue.CreateStringValue(jsonHeaders["Authorization"].GetString() + ApplicationData.Current.LocalSettings.Values["AccessToken"].ToString());

                    //Create the request object
                    RequestItem request = new RequestItem(requestObject["ApiUrl"].GetString(),
                                                       requestObject["Method"].GetString(),
                                                       jsonHeaders,
                                                       requestObject["Body"].GetObject());

                    //Create the data item object
                    DataItem item = new DataItem(itemObject["UniqueId"].GetString(),
                                                       itemObject["Title"].GetString(),
                                                       itemObject["Subtitle"].GetString(),
                                                       itemObject["ImagePath"].GetString()
                                                       );

                    // Add the request object to the item
                    item.Request = request;
                    
                    //Add the item to the group
                    group.Items.Add(item);


                }
                this.Groups.Add(group);
            }
        }
    }

    public class JsonObjectConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            // Format the json object as a correctly indented string
            JsonObject jsonObject = (JsonObject)value;
            return JsonConvert.SerializeObject(jsonObject, Formatting.Indented);
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            // Convert back from the string to a json object
            String strJson = (String)value;
            return JsonObject.Parse(strJson);
        }
    }
    public class MethodConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            // In the UI POST is true, Get is false
            return String.Compare((string)value, "POST") == 0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            // if the UI returns true, then method is POST, else it is GET
            return (bool)value ? "POST" : "GET";
        }
    }

}