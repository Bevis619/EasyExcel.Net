using System;

namespace EasyExcel.Extensions
{
    /// <summary>
    /// DateTime Extension Method
    /// </summary>
    public static class DateTimeExtension
    {
        /// <summary>
        /// Default DateTime Format
        /// </summary>
        public static string DefaultDateTimeFormat = "yyyy-MM-dd HH:mm:ss";

        /// <summary>
        /// Default Date Format
        /// </summary>
        public static string DefaultDateFormat = "yyyy-MM-dd";

        /// <summary>
        /// Convert to Default DateTime String
        /// </summary>
        /// <param name="datetime">datetime</param>
        /// <returns>default format string</returns>
        public static string ToDefaultString(this DateTime datetime) => datetime.ToString(DefaultDateTimeFormat);

        /// <summary>
        /// Convert to Default DateTime String
        /// </summary>
        /// <param name="datetime">datetime</param>
        /// <returns>default format string</returns>
        public static string ToDefaultDateString(this DateTime datetime) => datetime.ToString(DefaultDateFormat);
    }
}