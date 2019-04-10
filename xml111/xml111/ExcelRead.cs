using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace xml111
{
    class ExcelRead
    {
        static void Main3(string[] args)
        {
            IWorkbook workbook = null;  // 新建Iworkbook对象
            string fileName = @"C:\Users\CHAOCHEN\Desktop\A201811Original.xls";
            FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            if (fileName.IndexOf(".xls") > 0)  // 2003版本
            {
                workbook = new HSSFWorkbook(fileStream);
            }
            else if (fileName.IndexOf(".xlsx") > 0)
            {
                workbook = new XSSFWorkbook(fileStream);
            }
            ISheet sheet = workbook.GetSheetAt(0);  // 获取第一个工作表
            IRow row;  // 新建当前工作表行数据
            for(int i = 0; i < 30; i++)
            {
                row = sheet.GetRow(i);  // 读取第i行数据
                if(row != null)
                {
                    for(int j=0;j<row.LastCellNum; j++)  // 对工作表每一列
                    {
                        try
                        {
                            string cellValue = row.GetCell(j).ToString();  // 获取i行j列数据
                            Console.WriteLine(cellValue);
                        }
                        catch(Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }

                    }
                }
            }
            Console.ReadKey();
            fileStream.Close();
            workbook.Close();
        }
    }
}
