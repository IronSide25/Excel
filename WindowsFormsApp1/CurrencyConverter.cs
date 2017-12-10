using System.Web.Script.Serialization;
using System.Net;
using System.IO;
using System.Collections.Generic;
using System;

namespace WindowsFormsApp1
{
    internal class CurrencyConverter
    {
        // For example:
        //          "PLN"   0.23
        Dictionary <string, object> CurrencyTable;

        bool dataDownloaded;
        string exchangeRateDataUrl = "http://api.fixer.io/latest";

        public bool hasBeenConstructedInAProperWay()
        {
            return dataDownloaded;
        }

        public CurrencyConverter()
        {
            WebRequest request = WebRequest.Create(exchangeRateDataUrl);

            WebResponse response = request.GetResponse();

            // Handle a possible status error.
            if (((HttpWebResponse)response).StatusDescription != "OK")
            {
                // Should I handle it? xD     
                //MessageBox.Show("Failed to download exchange data. Check your internet connection.");
                dataDownloaded = false;
                return;
            }

            dataDownloaded = true;

            // Get the stream containing content returned by the server.
            Stream dataStream = response.GetResponseStream();
            // Open the stream using a StreamReader for easy access.
            StreamReader reader = new StreamReader(dataStream);
            // Read the content.
            string rawJson = reader.ReadToEnd();

            var jss = new JavaScriptSerializer();
            var table = jss.Deserialize<dynamic>(rawJson);

            CurrencyTable = (table["rates"]);

            reader.Close();
            response.Close();

        }

        // "PLN" - 0.23
        public float getRateOf(string currency)
        {
            if (currency.Equals("EUR"))
                return 1;
            double rate = (double)(decimal)CurrencyTable[currency];
            double reverseRate = Math.Round(1.0f / rate, 4);
            return (float)reverseRate;
        }


    }
}