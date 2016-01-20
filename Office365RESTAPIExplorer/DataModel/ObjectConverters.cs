// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

using Newtonsoft.Json;
using System;
using Windows.Data.Json;
using Windows.UI.Xaml.Data;

namespace Office365RESTExplorerforSites.Data
{
    /// <summary>
    /// Converts a JSON object to string and viceversa
    /// </summary>
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

    /// <summary>
    /// Converts a JSON string to a formatted string and viceversa.
    /// This class handles some exceptional cases where the string is not a well-formed
    /// JSON string.
    /// </summary>
    public class BodyConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            // Return an empty JSON object if the value is null
            if(value == null)
                return "{}";
            // Format the json string as a correctly indented string
            // otherwise return the raw string value
            JsonObject jsonObject;
            if (JsonObject.TryParse(value.ToString(), out jsonObject))
            {
                return JsonConvert.SerializeObject(jsonObject, Formatting.Indented);
            }
            else
            {
                return value.ToString();
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            // Return the string value
            return value.ToString();
        }
    }

    /// <summary>
    /// Converts a string value to boolean
    /// If the string is POST it converts to true, if it's GET converts to false
    /// </summary>
    public class MethodConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            // In the UI POST is true, Get is false
            return String.Compare((string)value, "POST", StringComparison.CurrentCultureIgnoreCase) == 0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            // if the UI returns true, then method is POST, else it is GET
            return (bool)value ? "POST" : "GET";
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