// This source code is based on Jiri Pik's FinancialDataForExcel add-on from here:
// http://www.assembla.com/spaces/FinancialDataForExcel/wiki

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Net;

using ExcelDna.Integration;

namespace MarketDataForExcel
{
    /// <summary>
    /// This class contains the actual code for obtaining market data from Yahoo! Finance
    /// </summary>
    public class YahooFinance
    {
        [ExcelFunction("Get latest FX Rate from Yahoo")]
        public static object GetFxRateFromYahooAsync([ExcelArgument("From Currency")] string fromCcy, [ExcelArgument("To Currency")] string toCcy)
        {
            return ExcelAsyncUtil.Run("GetFxRateFromYahooAsync", new[] {fromCcy, toCcy},
                () => GetFxRateFromYahoo(fromCcy, toCcy));
        }

        [ExcelFunction("Obtains historical market data from Yahoo")]
        public static object GetHistoricalDataFromYahooAsync(
                   [ExcelArgument("Yahoo Ticker")] string ticker,
                   [ExcelArgument("From Date")] DateTime fromDate,
                   [ExcelArgument("To Date")] DateTime toDate)
        {
            return ExcelAsyncUtil.Run("GetHistoricalDataFromYahooAsync", 
                new object[] { ticker, fromDate, toDate},
                () => GetHistoricalDataFromYahoo(ticker, fromDate, toDate));
        }

        /// <summary>
        /// This constant stores the duration of the cache in minutes. 
        /// </summary>
        private const double ExchangeRateCacheDuration = 5;

        /// <summary>
        /// This dictionary implements the caching of market data
        /// </summary>
        private static Dictionary<string, ExchangeRate> exchangeRateCache = new Dictionary<string, ExchangeRate>();

        /// <summary>
        /// This method retrieves the current FX Rate from Yahoo
        /// </summary>
        /// <param name="fromCcy">
        /// The FROM currency ISO code
        /// </param>
        /// <param name="toCcy">
        /// The TO currency ISO code
        /// </param>
        /// <returns>
        /// The method returns the current FX rate
        /// </returns>
        /// <exception cref="cellMatrixException">
        /// An exception occurs when the download of data fails for any reason
        /// </exception>
        [ExcelFunction("Get latest FX Rate from Yahoo")]
        public static double GetFxRateFromYahoo([ExcelArgument("From Currency")] string fromCcy, [ExcelArgument("To Currency")] string toCcy)
        {
            if (fromCcy.Equals(toCcy))
            {
                return 1.0;
            }

            if (fromCcy.Equals("TEST"))
            {
                throw new Exception("Test Successful");
            }

            var ticker = fromCcy.Trim().ToUpper() + toCcy.Trim().ToUpper() + "=X";

            if (exchangeRateCache.ContainsKey(ticker))
            {
                var exchangeRateCachedRecord = exchangeRateCache[ticker];
                var timeSpan = DateTime.Now.Subtract(exchangeRateCachedRecord.TimeStamp);
                if (timeSpan.TotalMinutes < ExchangeRateCacheDuration)
                {
                    return exchangeRateCachedRecord.Value;
                }

                exchangeRateCache.Remove(ticker);
            }

            var url = "http://download.finance.yahoo.com/d/quotes.csv?s=" + ticker + "'&f=l1&e=.cs";

            string exchangeRateInString;

            var webConnection = new WebClient();
            try
            {
                exchangeRateInString = webConnection.DownloadString(url);
            }
            catch (WebException ex)
            {
                throw new Exception("Unable to download the data! Check your Internet connection!", ex);
            }
            finally
            {
                webConnection.Dispose();
            }

            double exchangeRate;

            if (double.TryParse(exchangeRateInString, out exchangeRate))
            {
                exchangeRateCache.Add(ticker, new ExchangeRate { Ticker = ticker, TimeStamp = DateTime.Now, Value = exchangeRate });
                return exchangeRate;
            }

            throw new Exception("Returned value is not a number! Try later!");
        }

        /// <summary>
        /// This method retrieves a timeseries of historical stock data
        /// </summary>
        /// <param name="ticker">
        /// The Yahoo! ticker.
        /// </param>
        /// <param name="fromDate">
        /// The FROM date.
        /// </param>
        /// <param name="toDate">
        /// The TO date.
        /// </param>
        /// <returns>
        /// The method returns an array of historical data.
        /// </returns>
        /// <exception cref="cellMatrixException">
        /// An exception is thrown if for any reason the download fails.
        /// </exception>
        [ExcelFunction("Obtains historical market data from Yahoo")]
        public static object[,] GetHistoricalDataFromYahoo(
                   [ExcelArgument("Yahoo Ticker")] string ticker,
                   [ExcelArgument("From Date")] DateTime fromDate,
                   [ExcelArgument("To Date")] DateTime toDate)
        {
            var begin = fromDate;
            var end = toDate;

            var yahooURL =
               @"http://ichart.finance.yahoo.com/table.csv?s=" +
               ticker + @"&a=" + (begin.Month - 1).ToString(CultureInfo.InvariantCulture) + @"&b=" + begin.Day.ToString(CultureInfo.InvariantCulture) +
               @"&c=" + begin.Year.ToString(CultureInfo.InvariantCulture) + @"&d=" + (end.Month - 1).ToString(CultureInfo.InvariantCulture) + @"&e=" + end.Day.ToString(CultureInfo.InvariantCulture) + @"&f=" + end.Year.ToString(CultureInfo.InvariantCulture) +
               @"&g=d&ignore=.csv";

            string historicalData;
            var webConnection = new WebClient();
            try
            {
                historicalData = webConnection.DownloadString(yahooURL);
            }
            catch (WebException ex)
            {
                throw new Exception("Unable to download the data! Check your Internet Connection!", ex);
            }
            finally
            {
                webConnection.Dispose();
            }

            historicalData = historicalData.Replace("\r", string.Empty);
            var rows = historicalData.Split('\n');
            var headings = rows[0].Split(',');

            var excelData = new object[rows.Length + 1, headings.Length];
            for (var i = 0; i < headings.Length; ++i)
            {
                excelData[0, i] = headings[i];
            }

            for (var i = 1; i < rows.Length; ++i)
            {
                var thisRow = rows[i].Split(',');
                if (thisRow.Length == headings.Length)
                {
                    excelData[i, 0] = DateTime.Parse(thisRow[0]);
                    for (var j = 1; j < headings.Length; ++j)
                    {
                        excelData[i, j] = double.Parse(thisRow[j]);
                    }
                }
            }

            return excelData;
        }

        /// <summary>
        /// This class is used for caching of the retrieved results.
        /// </summary>
        private class ExchangeRate
        {
            /// <summary>
            /// Gets or sets the Yahoo! Ticker
            /// </summary>
            public string Ticker { get; set; }

            /// <summary>
            /// Gets or sets the Exchange Rate
            /// </summary>
            public double Value { get; set; }

            /// <summary>
            /// Gets or sets TimeStamp of the value
            /// </summary>
            public DateTime TimeStamp { get; set; }
        }
    }
}