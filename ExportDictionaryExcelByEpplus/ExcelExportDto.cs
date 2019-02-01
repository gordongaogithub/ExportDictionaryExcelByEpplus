using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportDictionaryExcelByEpplus
{
    /// <summary>
    /// 导出类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExcelExportDto<T>
    {
        public ExcelExportDto(string columnName, Func<T, object> columnValue)
        {
            ColumnName = columnName;
            ColumnValue = columnValue;
        }
        public string ColumnName { get; set; }

        public Func<T, object> ColumnValue { get; set; }
    }
}
