using System.ComponentModel;
using System.Data;
using System.Reflection;

namespace EwatchIntegrateSystem.Methods
{
    /// <summary>
    /// DataTable To List
    /// </summary>
    public static class DataTableToList
    {
        /// <summary>
        /// To List Function
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public static IList<T> ToList<T>(this DataTable data) where T : new()
        {
            Type type = typeof(T);
            IList<PropertyInfo> properties = typeof(T).GetProperties().ToList();
            IList<T> result = new List<T>();

            foreach (DataRow row in data.Rows)
            {
                var item = CreateItemFromRow<T>(row, properties);
                result.Add(item);
            }
            return result;
        }
        /// <summary>
        /// Create Item From Row
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="row"></param>
        /// <param name="properties"></param>
        /// <returns></returns>
        private static T CreateItemFromRow<T>(DataRow row, IList<PropertyInfo> properties) where T : new()
        {
            T item = new T();
            foreach (var property in properties)
            {
                string? value = row[property.Name].ToString() == "" ? "0" : row[property.Name].ToString();
                property.SetValue(item,value, null);
            }
            return item;
        }
    }
}
