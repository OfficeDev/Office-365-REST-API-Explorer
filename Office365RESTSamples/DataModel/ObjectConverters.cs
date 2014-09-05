using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Newtonsoft.Json;
using Windows.UI.Xaml.Data;
using Windows.Data.Json;

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
