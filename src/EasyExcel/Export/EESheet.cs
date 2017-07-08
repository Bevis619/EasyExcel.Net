using System.Collections.Generic;

namespace EasyExcel.Export
{
    /// <summary>
    /// export excel sheet info 
    /// </summary>
    public class EESheet
    {
        /// <summary>
        /// sheet name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// sheet data collection
        /// </summary>
        public IEnumerable<object> Sheets { get; set; }

        /// <summary>
        /// ctor
        /// </summary>
        /// <param name="sheets">sheet data collection</param>
        /// <param name="name">sheet name</param>
        public EESheet(IEnumerable<object> sheets, string name = "")
        {
            this.Sheets = sheets;
            this.Name = name;
        }
    }
}