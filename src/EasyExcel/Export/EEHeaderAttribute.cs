using System;

namespace EasyExcel.Export
{
    /// <summary>
    /// easy export to excel header attribute
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class EEHeaderAttribute : Attribute
    {
        /// <summary>
        /// Header Name
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Place Sequence
        /// </summary>
        public uint Sequence { get; set; }

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="name">Header Name</param>
        /// <param name="sequence">Place Sequence</param>
        public EEHeaderAttribute(string name, uint sequence = 0)
        {
            this.Name = name;
            this.Sequence = sequence;
        }
    }
}