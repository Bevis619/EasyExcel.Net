using EasyExcel.Export;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;

namespace EasyExcel.Extensions
{
    /// <summary>
    /// IEnumerable Extension Method
    /// </summary>
    public static class IEnumerableExtension
    {
        /// <summary>
        /// convert to datatable
        /// </summary>
        /// <param name="models">enumerable data</param>
        /// <returns>datatable</returns>
        public static DataTable ToDataTable(this IEnumerable<object> models)
        {
            if (null == models) throw new ArgumentNullException("models is null");
            if (!models.Any()) throw new ArgumentException("models is empty");
            var table = new DataTable();
            var type = models.First().GetType();
            if (type.IsDefined(typeof(EESheetAttribute)))
                table.TableName = type.GetCustomAttribute<EESheetAttribute>(true).Name;
            var hash = new Dictionary<string, string>();
            var attributes = new List<EEHeaderAttribute>();
            foreach (var prop in type.GetProperties(BindingFlags.Instance | BindingFlags.Public))
            {
                if (!prop.IsDefined(typeof(EEHeaderAttribute))) continue;
                var attr1 = prop.GetCustomAttribute<EEHeaderAttribute>();
                attributes.Add(attr1);
                hash.Add(prop.Name, attr1.Name);
            }

            foreach (var item in attributes.OrderBy(p => p.Sequence))
            {
                table.Columns.Add(item.Name);
            }

            DataRow row = null;
            foreach (var model in models)
            {
                row = table.NewRow();
                foreach (var property in model.GetType().GetProperties(BindingFlags.Instance | BindingFlags.Public))
                {
                    if (!hash.ContainsKey(property.Name)) continue;
                    row[hash[property.Name]] = property.GetValue(model);
                }

                table.Rows.Add(row);
            }

            return table;
        }
    }
}