using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportDictionaryExcelByEpplus
{
    /// <summary>
    /// 实体
    /// </summary>
    public class Student
    {
        public String Name { get; set; }

        public String Code { get; set; }

        public Dictionary<string, string> Dictionarys { get; set; }
    }
}
