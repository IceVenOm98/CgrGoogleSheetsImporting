using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Net;
using System.Reflection;

namespace CgrGoogleSheetsImporting
{
    public class DataProcessor<T> where T : new()
    {
        private Dictionary<string, string> AttributesFields;
        readonly private T Entity;
        readonly private string SpreadsheetId;
        readonly private string SheetName;
        public IApiConnector ApiConnector { get; set; }

        /// <summary>
        /// Конструктор обработчика данных
        /// </summary>
        /// <param name="entity">Сущность типа, список которого необходимо получить</param>
        /// <param name="spreadsheetId">Id таблицы</param>
        /// <param name="sheetName">Название листа</param>
        /// <param name="apiConnector">Коннектор API</param>
        public DataProcessor(T entity, string spreadsheetId, string sheetName, IApiConnector apiConnector)
        {
            Entity = entity;
            SpreadsheetId = spreadsheetId;
            SheetName = sheetName;
            ApiConnector = apiConnector;
            CreateAttributesFieldsDictionary();
        }

        /// <summary>
        /// Создает словарь пар атрибутов и полей
        /// </summary>
        private void CreateAttributesFieldsDictionary()
        {
            AttributesFields = new Dictionary<string, string>();
            foreach (PropertyInfo pi in Entity.GetType().GetProperties())
            {
                string fieldName = pi.Name;
                string fieldAttribute = Entity.GetType().GetProperty(fieldName).GetCustomAttribute<DisplayAttribute>().Name;
                AttributesFields.Add(fieldAttribute, fieldName);
            }
        }

        /// <summary>
        /// Создает список сущностей
        /// </summary>
        /// <returns>
        /// Список объектов типа T
        /// </returns>
        public List<T> GetEntityList()
        {
            SheetData sheetData = ApiConnector.GetSheetData(SpreadsheetId, SheetName);
            List<String> columns = sheetData.Values[0];
            sheetData.Values.RemoveAt(0);
            List<T> list = new List<T>();
            int rowLength = columns.Count;
            foreach (List<string> row in sheetData.Values)
            {
                AddEmptyCells(row, rowLength);
                if (row[0] == "")
                {
                    break;
                }
                list.Add(CreateEntity(row, columns));
            }
            return list;
        }

        /// <summary>
        /// Создает сущность на основе строки из таблицы. Типизирует значения полей.
        /// </summary>
        /// <param name="row">Строка значений (полей) сущности</param>
        /// <param name="columns">Заголовки полей</param>
        /// <returns>Сущность T</returns>
        private T CreateEntity(List<string> row, List<string> columns)
        {
            T entity = new T();
            for (int columnNumber = 0; columnNumber < row.Count; columnNumber++)
            {
                string fieldName = AttributesFields[columns[columnNumber]]; //получение имени свойства по названию колонки таблицы
                SetField(entity, fieldName, row[columnNumber]);
            }
            return entity;
        }

        private void SetField(T entity, string fieldName, string value)
        {
            Type type = Entity.GetType().GetProperty(fieldName).PropertyType;
            if (!TryChangeType(value, type, out object typedValue))
            {
                typedValue = Entity.GetType().GetProperty(fieldName).GetValue(entity);
            }
            Entity.GetType().GetProperty(fieldName).SetValue(entity, typedValue);
        }

        /// <summary>
        /// Добавление пустых ячеек для заполнения таблицы
        /// </summary>
        /// <param name="row">Строка с пустыми значениями</param>
        /// <param name="rowLength">Длина строки</param>
        private void AddEmptyCells(List<string> row, int rowLength)
        {
            for (int i = row.Count; i < rowLength; i++)
            {
                row.Add(string.Empty);
            }
        }

        /// <summary>
        /// Преобразует объект в передаваемый тип.
        /// </summary>
        /// <param name="value">Объект, который нужно преобразовать.</param>
        /// <param name="conversionType">Тип, в который нужно преобразовать объект.</param>
        /// <param name="typedValue">Преобразованный объект.</param>
        /// <returns>Значение true, если объект успешно преобразован; в противном случае — значение false.</returns>
        private bool TryChangeType(Object value, Type conversionType, out Object typedValue)
        {
            typedValue = null;
            if (conversionType == null || value == null || value.ToString()==string.Empty)
            {
                return false;
            }
            if (!(value is IConvertible))
            {
                if (value.GetType() == conversionType)
                {
                    typedValue = value;
                    return true;
                }
                return false;
            }
            typedValue = Convert.ChangeType(value, conversionType);
            return true;
        }
    }
}
