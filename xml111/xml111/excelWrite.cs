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
    class excelWrite
    {
        static void Main2(string[] args)
        {
            HSSFWorkbook workbook2003 = new HSSFWorkbook();
            workbook2003.CreateSheet("Sheet1");
            HSSFSheet SheetOne = (HSSFSheet)workbook2003.GetSheet("Sheet1");  // 获取名称为sheet1的工作表
            // 对工作表先添加行，下标从0开始。
            for(int i = 0; i < 10; i++)
            {
                SheetOne.CreateRow(i);
            }
            // 对每一行创建10个单元格
            HSSFRow SheetRow = (HSSFRow)SheetOne.GetRow(0);  // 获取sheet1工作表的首行
            HSSFCell[] SheetCell = new HSSFCell[10];  // 每一行10个单元格
            for (int i = 0; i < 10; i++)
            {
                SheetCell[i] = (HSSFCell)SheetRow.CreateCell(i);  // 为第一行创建10个单元格
            }
            // 创建之后就可以赋值了
            SheetCell[0].SetCellValue(true);
            SheetCell[1].SetCellValue(0.00001);
            SheetCell[2].SetCellValue("Excel2003");
            SheetCell[3].SetCellValue("123455644446454545");

            for(int i = 4; i < 10; i++)
            {
                SheetCell[i].SetCellValue(i);  // 循环赋值为整型
            }
            FileStream file2003 = new FileStream(@"C:\Users\CHAOCHEN\Desktop\test001.xls", FileMode.Create);
            workbook2003.Write(file2003);
            file2003.Close();
            workbook2003.Close();
        }
    }
}
