using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace CgrGoogleSheetsImporting
{
    /// <summary>
    /// Коннектор для получения данных из гугл-таблицы
    /// </summary>
    public class GoogleSheetsApiConnector : IApiConnector
    {
        private readonly Uri ApiUrl;
        readonly string Key;

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
            try
            {
                var webClient = new WebClient
                {
                    Encoding = System.Text.Encoding.UTF8
                };
                Uri requestUrl = new Uri(string.Format("{0}{1}?key={2}", ApiUrl.OriginalString, spreadsheetId, Key));
                var response = webClient.DownloadString(requestUrl);
                return JsonConvert.DeserializeObject<SpreadSheetsData>(response);
            }
            catch (Exception)
            {
                throw;
            }
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
                var response = webClient.DownloadString(requestUrl);
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
            var webClient = new WebClient
            {
                Encoding = System.Text.Encoding.UTF8
            };
            try
            {
                var response = webClient.DownloadString(requestUrl);
                return JsonConvert.DeserializeObject<SheetData>(response);
            }
            catch (WebException e)
            {
                throw new WrongDownloadingDataException(e);
            }
            catch (JsonSerializationException e)
            {
                throw new JsonSerializationException("Ошибка сериализации объектов из таблицы. Проверьте структуру объектов и название атрибутов полей", e);
            }
            catch (Exception e)
            {
                throw new Exception("Ошибка получения информации", e);
            }
        }
    }
}
