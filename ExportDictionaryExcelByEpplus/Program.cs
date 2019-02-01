using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExportDictionaryExcelByEpplus
{
    class Program
    {
        static void Main(string[] args)
        {
            //获得数据
            List<Student> studentList = new List<Student>();
            for (int i = 0; i < 10; i++)
            {
                Student s = new Student();
                s.Code = "c" + i;
                s.Name = "n" + i;
                studentList.Add(s);
            }

            //获得不固定数据
            for (int i = 0; i < studentList.Count; i++)
            {
                Dictionary<string, string> dictionarys = new Dictionary<string, string>();
                dictionarys.Add("D1", "d1" + i);
                dictionarys.Add("D2", "d2" + i);
                studentList[i].Dictionarys = dictionarys;
            }


            //创建excel
            string fileName = @"d:\" + "导出excel" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            FileInfo newFile = new FileInfo(fileName);
            using (ExcelPackage package = new ExcelPackage(newFile))
            {
                #region 固定列
                List<ExcelExportDto<Student>> excelExportDtoList = new List<ExcelExportDto<Student>>();
                excelExportDtoList.Add(new ExcelExportDto<Student>("Code", _ => _.Code));
                excelExportDtoList.Add(new ExcelExportDto<Student>("Name", _ => _.Name));

                List<string> columnsNameList = new List<string>();
                List<Func<Student, object>> columnsValueList = new List<Func<Student, object>>();
                foreach (var item in excelExportDtoList)
                {
                    columnsNameList.Add(item.ColumnName);
                    columnsValueList.Add(item.ColumnValue);
                }

                #endregion

                #region 不固定列
                List<ExcelExportDto<Dictionary<string, string>>> excelExportDictionaryDtoList = new List<ExcelExportDto<Dictionary<string, string>>>();
                List<string> columnsNameDictionaryList = new List<string>();
                List<string> dictionaryKeys = EpplusHelper.GetDictionaryKeys(studentList[0].Dictionarys);

                if (studentList.Count > 0)
                {
                    for (int i = 0; i < dictionaryKeys.Count; i++)
                    {
                        var index = i;
                        excelExportDictionaryDtoList.Add(new ExcelExportDto<Dictionary<string, string>>(dictionaryKeys[i], _ => _.FirstOrDefault(q => q.Key == dictionaryKeys[i]).Value));
                    }
                    foreach (var item in excelExportDictionaryDtoList)
                    {
                        columnsNameDictionaryList.Add(item.ColumnName);
                    }
                }
                #endregion

                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Test");
                worksheet.OutLineApplyStyle = true;
                //添加表头
                EpplusHelper.AddHeader(worksheet, columnsNameList.ToArray());
                //添加数据
                EpplusHelper.AddObjects(worksheet, 2, studentList, columnsValueList.ToArray());
                if (studentList.Count > 0)
                {
                    //添加动态表头
                    EpplusHelper.AddHeader(worksheet, columnsNameList.ToArray(), columnsNameDictionaryList.ToArray());
                    //添加动态数据
                    EpplusHelper.AddObjects(worksheet, 2, studentList, columnsValueList.ToArray(), dictionaryKeys);
                }
                package.Save();
            }
        }
    }
}
