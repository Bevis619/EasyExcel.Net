using System;

namespace EasyExcel.Export
{
    /// <summary>
    /// easy export to excel sheet attribute
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class EESheetAttribute : Attribute
    {
        /// <summary>
        /// sheet name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="name">sheet name</param>
        public EESheetAttribute(string name)
        {
            this.Name = name;
        }
    }
}