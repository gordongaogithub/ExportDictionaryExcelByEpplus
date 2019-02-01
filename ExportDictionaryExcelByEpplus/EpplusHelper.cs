using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportDictionaryExcelByEpplus
{
    public static class EpplusHelper
    {


        /// <summary>
        /// 添加表头
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="headerTexts"></param>
        public static void AddHeader(ExcelWorksheet sheet, params string[] headerTexts)
        {
            for (var i = 0; i < headerTexts.Length; i++)
            {
                AddHeader(sheet, i + 1, headerTexts[i]);
            }
        }

        /// <summary>
        /// 添加表头
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columnIndex"></param>
        /// <param name="headerText"></param>
        public static void AddHeader(ExcelWorksheet sheet, int columnIndex, string headerText)
        {
            sheet.Cells[1, columnIndex].Value = headerText;
            sheet.Cells[1, columnIndex].Style.Font.Bold = true;
        }

        /// <summary>
        /// 添加数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="items"></param>
        /// <param name="propertySelectors"></param>
        public  static void AddObjects(ExcelWorksheet sheet, int startRowIndex, IList<Student> items, Func<Student, object>[] propertySelectors)
        {
            for (var i = 0; i < items.Count; i++)
            {
                for (var j = 0; j < propertySelectors.Length; j++)
                {
                    sheet.Cells[i + startRowIndex, j + 1].Value = propertySelectors[j](items[i]);
                }
            }
        }

        /// <summary>
        /// 添加动态表头
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="headerTexts"></param>
        /// <param name="headerTextsDictionary"></param>
        public  static void AddHeader(ExcelWorksheet sheet, string[] headerTexts, string[] headerTextsDictionary)
        {                      
            for (var i = 0; i < headerTextsDictionary.Length; i++)
            {
                AddHeader(sheet, i + 1 + headerTexts.Length, headerTextsDictionary[i]);
            }
        }

        /// <summary>
        /// 添加动态数据
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="startRowIndex"></param>
        /// <param name="items"></param>
        /// <param name="propertySelectors"></param>
        /// <param name="dictionaryKeys"></param>

        public static void AddObjects(ExcelWorksheet sheet, int startRowIndex, IList<Student> items, Func<Student, object>[] propertySelectors, List<string> dictionaryKeys)
        {           
            for (var i = 0; i < items.Count; i++)
            {
                for (var j = 0; j < dictionaryKeys.Count; j++)
                {
                    sheet.Cells[i + startRowIndex, j + 1 + propertySelectors.Length].Value = items[i].Dictionarys[dictionaryKeys[j]];
                }
            }
                    
        }


        public static List<String> GetDictionaryKeys(Dictionary<string, string> dics)
        {
            List<string> resultList = new List<string>();
            foreach (KeyValuePair<string, string> kvp in dics)
            {
                resultList.Add(kvp.Key);
            }
            return resultList;
        }

    }
}
