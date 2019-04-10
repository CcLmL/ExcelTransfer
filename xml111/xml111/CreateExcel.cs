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
    class CreateExcel
    {
        static void Main1(string[] args)
        {
            HSSFWorkbook workbook2003 = new HSSFWorkbook();  // 新建xls工作簿
            workbook2003.CreateSheet("Sheet1");
            workbook2003.CreateSheet("Sheet2");
            workbook2003.CreateSheet("Sheet3");
            FileStream file2003 = new FileStream(@"C:\Users\CHAOCHEN\Desktop\C-Daily\test001.xls", FileMode.Create);  // 创建文件读写流
            workbook2003.Write(file2003);
            file2003.Close();
            workbook2003.Close();

            XSSFWorkbook workbook2007 = new XSSFWorkbook();
            workbook2007.CreateSheet("Sheet1");
            workbook2007.CreateSheet("Sheet2");
            workbook2007.CreateSheet("Sheet3");
            FileStream file2007 = new FileStream(@"C:\Users\CHAOCHEN\Desktop\C-Daily\test002.xlsx", FileMode.Create);
            workbook2007.Write(file2007);
            workbook2007.Close();
            file2007.Close();
        }
    }
}
