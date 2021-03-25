using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using CgrGoogleSheetsImporting.Exceptions;
using System.IO;

namespace CgrGoogleSheetsImporting
{
    /// <summary>
    /// Коннектор для получения данных из гугл-таблицы
    /// </summary>
    public class GoogleSheetsApiConnector : IApiConnector
    {
        private readonly Uri ApiUrl;
        readonly string Key;
        public bool Status { get; }
        public string StatusMessage { get; }


        /// <summary>
        /// Конструктор коннектора
        /// </summary>
        /// <param name="apiUrl">Адрес API сервиса</param>
        /// <param name="key">Ключ доступа к Google Sheets API (необходим проект Google Cloud Platform)</param>
        public GoogleSheetsApiConnector(string apiUrl,  string key)
        {
            ApiUrl = new Uri(apiUrl);
            Key = key;
        }


        /// <summary>
        /// Возвращает информацию о гугл-таблице
        /// </summary>
        /// <param name="spreadsheetId">Идентификатор таблицы</param>
        /// <returns>SheetData (Лист таблицы)</returns>
        ///
        public SpreadSheetsData GetSpreadSheetsData(string spreadsheetId)
        {
            Uri requestUrl = new Uri(string.Format("{0}{1}?key={2}", ApiUrl.OriginalString, spreadsheetId, Key));
            TryDownloadString(requestUrl, out string response);
            return JsonConvert.DeserializeObject<SpreadSheetsData>(response);
        }

        /// <summary>
        /// Возвращает информацию о первом листе гугл-таблицы
        /// </summary>
        /// <param name="spreadsheetId">Идентификатор таблицы</param>
        /// <returns>SheetData (Лист таблицы)</returns>
        ///
        public SheetData GetSheetData(string spreadsheetId)
        {
            string sheetName = GetSpreadSheetsData(spreadsheetId).Sheets[0].Properties.Title;
            try
            {
                var webClient = new WebClient
                {
                    Encoding = System.Text.Encoding.UTF8
                };
                Uri requestUrl = new Uri(string.Format("{0}{1}/values/{2}?key={3}", ApiUrl.OriginalString, spreadsheetId, sheetName, Key));
                TryDownloadString(requestUrl,  out string response);
                return JsonConvert.DeserializeObject<SheetData>(response);
            }
            catch (Exception)
            {
                throw;
            }
        }

        /// <summary>
        /// Возвращает информацию о листе гугл-таблицы по названию
        /// </summary>
        /// <param name="spreadsheetId">Идентификатор таблицы</param>
        /// <param name="sheetName">Название листа таблицы</param>
        /// <returns>SheetData (Лист таблицы)</returns>
        public SheetData GetSheetData(string spreadsheetId, string sheetName)
        {
            Uri requestUrl = new Uri(string.Format("{0}{1}/values/{2}?key={3}", ApiUrl.OriginalString, spreadsheetId, sheetName, Key));
            TryDownloadString(requestUrl, out string response);
            return JsonConvert.DeserializeObject<SheetData>(response);

        }


        private bool TryDownloadString(Uri requestUrl, out string response)
        {
            response = null;
            var webClient = new WebClient
            {
                Encoding = System.Text.Encoding.UTF8
            };
            try
            {
                response = webClient.DownloadString(requestUrl);
                return true;
            }
            catch (WebException e)
            {
                var error = new StreamReader(e.Response.GetResponseStream()).ReadToEnd();
                error = error.Substring(13, error.Length - 16);
                ApiError apiError = JsonConvert.DeserializeObject<ApiError>(error);
                if (e.Status == WebExceptionStatus.ProtocolError)
                {
                    throw new NotValidRequestException(e, apiError.Message);
                }
                else
                {
                    throw;
                }
            }
        }
    }
}
