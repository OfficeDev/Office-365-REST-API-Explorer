/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

using Office365RESTExplorerforSites.Helpers;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Windows.Data.Json;
using Windows.Storage;

// The data model represents a hierarchical organization of objects as follows
// DataSource -> DataGroups -> DataItems -> ResponseItem
//                                       -> RequestItem
// All the objects are read from the InitialData.json file except for the ResponseItem object.
// ResponseItem object is created by issuing a request to the REST endpoint in the GetResponseAsync method
// DataGroups are used to populate the page that shows the SharePoint objects (lists, list items, documents)
// DataItems are used to populate the page that has the REST request and response. Usually, these are the CRUD operations.

namespace Office365RESTExplorerforSites.Data
{
    /// <summary>
    /// A data source object that models an HTTP response. Note that this object doesn't have a representation in the InitialData.json file
    /// </summary>
    public class ResponseItem
    {
        public ResponseItem(Uri responseUri, string status, JsonObject headers, string body)
        {
            this.ResponseUri = responseUri;
            this.Status = status;
            this.Headers = headers;
            this.Body = body;
        }

        public Uri ResponseUri { get; private set; }
        public JsonObject Headers { get; private set; }
        public string Body { get; private set; }
        public string Status { get; private set; }
    }

    /// <summary>
    /// A data source object that models an HTTP request. The InitialData.json file has one RequestItem object for eevery DataItem
    /// </summary>
    public class RequestItem : INotifyPropertyChanged
    {
        public RequestItem(Uri apiUrl, string method, JsonObject headers, string body)
        {
            this.ApiUrl = apiUrl.OriginalString;

            // Validate that the method is either "GET" or "POST"
            if (String.Compare(method, "GET", StringComparison.CurrentCultureIgnoreCase) != 0 
                && String.Compare(method, "POST", StringComparison.CurrentCultureIgnoreCase) != 0)
                throw new ArgumentOutOfRangeException("method", "The HTTP method can only be GET or POST.");
            else
                this.Method = method;

            this.Headers = headers;
            this.Body = body;
        }

        public string ApiUrl { get; set; }
        public JsonObject Headers { get; set; }
        public string Body { get; set; }
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
    /// Item data model. An item one of the CRUD operations that appear in the InitialData.json file
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
    /// Group data model. A group is each one of the SharePoint elements, for now those are lists, list items, and documents.
    /// Every group has items in it, usually CRUD operations.
    /// </summary>
    public class DataGroup
    {
        public DataGroup(String uniqueId, String title, String subtitle, String imagePath, String moreInfoText, Uri moreInfoUri)
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
        public Uri MoreInfoUri { get; private set; }
        public string MoreInfoText { get; private set; }
        public ObservableCollection<DataItem> Items { get; private set; }

        public override string ToString()
        {
            return this.Title;
        }
    }

    /// <summary>
    /// Creates a collection of groups and items with content read from a static json file.
    /// DataSource initializes with data read from a static json file included in the 
    /// project.  This provides sample data at both design-time and run-time.
    /// It also provides a method to issues a request to a REST endpoint and get a response object
    /// </summary>
    public sealed class DataSource
    {
        private static DataSource _dataSource = new DataSource();

        private ObservableCollection<DataGroup> _groups = new ObservableCollection<DataGroup>();

        public ObservableCollection<DataGroup> Groups
        {
            get { return this._groups; }
        }

        public static async Task<IEnumerable<DataGroup>> GetGroupsAsync()
        {
            await _dataSource.GetSampleDataAsync();

            return _dataSource.Groups;
        }

        public static void Clear()
        {
            _dataSource.Groups.Clear();
        }

        public static async Task<DataGroup> GetGroupAsync(string uniqueId)
        {
            await _dataSource.GetSampleDataAsync();
            // Simple linear search is acceptable for small data sets
            var matches = _dataSource.Groups.Where((group) => group.UniqueId.Equals(uniqueId));
            if (matches.Count() == 1) return matches.First();
            return null;
        }

        public static async Task<DataItem> GetItemAsync(string uniqueId)
        {
            await _dataSource.GetSampleDataAsync();
            // Simple linear search is acceptable for small data sets
            var matches = _dataSource.Groups.SelectMany(group => group.Items).Where((item) => item.UniqueId.Equals(uniqueId));
            if (matches.Count() == 1) return matches.First();
            return null;
        }

        public static async Task<ResponseItem> GetResponseAsync(RequestItem request)
        {
            HttpWebRequest endpointRequest;

            //Validate that the resulting URI is well-formed.
            Uri endpointUri = new Uri(new Uri(AuthenticationHelper.ServiceResourceId), request.ApiUrl);

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
                // If the request is a create or update file operation, we use the rawBody parameter
                // otherwise we use the request.Body
                string postData = request.Body;
                UTF8Encoding encoding = new UTF8Encoding();
                byte[] byte1 = encoding.GetBytes(postData);
                System.IO.Stream newStream = await endpointRequest.GetRequestStreamAsync();
                newStream.Write(byte1, 0, byte1.Length);
            }

            Stream responseStream;
            WebHeaderCollection responseHeaders;
            string status;
            Uri responseUri;
            JsonObject headers = null;
            string body = null;
            string responseString = string.Empty;

            try
            {
                // If the request is successful we can use the endpointResponse object
                HttpWebResponse endpointResponse = (HttpWebResponse)await endpointRequest.GetResponseAsync();
                status = (int)endpointResponse.StatusCode + " - " + endpointResponse.StatusDescription;
                responseStream = endpointResponse.GetResponseStream();
                responseUri = endpointResponse.ResponseUri;
                responseHeaders = endpointResponse.Headers;
            }
            catch (WebException we)
            {
                // If the request fails, we must use the response stream from the exception
                status = we.Message;
                responseStream = we.Response.GetResponseStream();
                responseUri = we.Response.ResponseUri;
                responseHeaders = we.Response.Headers;
            }

            using (StreamReader reader = new StreamReader(responseStream, Encoding.UTF8))
            {
                responseString = await reader.ReadToEndAsync();
            }

            // Free resources used by the stream
            responseStream.Dispose();

            if (!String.IsNullOrEmpty(responseString))
            {
                body = responseString;
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
            Uri dataUri = new Uri("ms-appx:///DataModel/InitialData.json"); 
            StorageFile file = await StorageFile.GetFileFromApplicationUriAsync(dataUri);
            string jsonText = await FileIO.ReadTextAsync(file);

            // Ensure that we have a valid access token before updating the data
            string accessToken = await AuthenticationHelper.EnsureAccessTokenAvailableAsync();

            lock (this.Groups)
            {
                if (this.Groups.Count != 0)
                    return;

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
                                                                new Uri(groupObject["MoreInfoUri"].GetString()));

                    foreach (JsonValue itemValue in groupObject["Items"].GetArray())
                    {
                        JsonObject itemObject = itemValue.GetObject();
                        JsonObject requestObject = itemObject["Request"].GetObject();

                        //Add the Authorization header with the access token.
                        JsonObject jsonHeaders = requestObject["Headers"].GetObject();
                        jsonHeaders["Authorization"] = JsonValue.CreateStringValue(jsonHeaders["Authorization"].GetString()
                            + accessToken);

                        // The body can be a JSON object or string, we need to 
                        // determine the type of JSON value and use the right 
                        // method to get the value.
                        string strBody;
                        if (requestObject["Body"].ValueType == JsonValueType.Object)
                        {
                            strBody = requestObject["Body"].GetObject().Stringify();
                        }
                        else if (requestObject["Body"].ValueType == JsonValueType.String)
                        {
                            strBody = requestObject["Body"].GetString();
                        }
                        else
                        {
                            throw new NotSupportedException("The body should only be of value JSON object or JSON string.");
                        }

                        //Create the request object
                        RequestItem request = new RequestItem(new Uri(requestObject["ApiUrl"].GetString(), UriKind.Relative),
                                                           requestObject["Method"].GetString(),
                                                           jsonHeaders,
                                                           strBody);

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
    }
}
