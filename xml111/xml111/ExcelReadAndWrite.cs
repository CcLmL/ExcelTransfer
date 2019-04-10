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
    class ExcelReadAndWrite
    {
        static void Main1(string[] args)
        {
            // 创建读取的对象
            IWorkbook workbookRead = null;
            string filenameRead = @"C:\Users\CHAOCHEN\Desktop\test001.xls";
            FileStream filestreamRead = new FileStream(filenameRead, FileMode.Open, FileAccess.Read);

            // 创建写入的excel对象
            HSSFWorkbook workbookWrite = new HSSFWorkbook();
            workbookWrite.CreateSheet("Sheet1");
            HSSFSheet SheetOne = (HSSFSheet)workbookWrite.GetSheet("Sheet1");

            if (filenameRead.Contains(".xls"))
            {
                workbookRead = new HSSFWorkbook(filestreamRead);
            }
            else if (filenameRead.Contains(".xlsx"))
            {
                workbookRead = new XSSFWorkbook(filestreamRead);
            }

            ISheet sheet1 = workbookRead.GetSheetAt(0);
            IRow row;
            HSSFRow Row;
            HSSFCell[] sheetCell = new HSSFCell[10];
            for (int i = 0; i < sheet1.LastRowNum; i++)
            {
                row = sheet1.GetRow(i);
                SheetOne.CreateRow(i);
                Row = (HSSFRow)SheetOne.GetRow(i);
                if (row != null)
                {
                    for(int j = 0; j < row.LastCellNum; j++)
                    {
                        string cellValue = row.GetCell(j).ToString();
                        sheetCell[j] = (HSSFCell)Row.CreateCell(j);
                        sheetCell[j].SetCellValue(cellValue + "cc");
                    }
                }
            }
            FileStream fileStreamWrite = new FileStream(@"C:\Users\CHAOCHEN\Desktop\test.xls", FileMode.Create);
            workbookWrite.Write(fileStreamWrite);
            fileStreamWrite.Close();
            filestreamRead.Close();
            workbookWrite.Close();
            workbookRead.Close();
        }
    }
}
