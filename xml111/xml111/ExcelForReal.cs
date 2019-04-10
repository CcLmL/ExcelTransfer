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
    class ExcelFoReal
    {
        static void Main11(string[] args)
        {
            // 创建读取的对象
            IWorkbook workbookRead = null;
            //string filenameRead = @"C:\Users\CHAOCHEN\Desktop\A201811Original.xls";
            Console.WriteLine("请输入文件路径：");
            string filenameRead = Console.ReadLine();
            FileStream filestreamRead = new FileStream(filenameRead, FileMode.Open, FileAccess.Read);

            // 创建写入的excel对象
            HSSFWorkbook workbookWrite = new HSSFWorkbook();
            workbookWrite.CreateSheet("Sheet1");
            workbookWrite.CreateSheet("Sheet2");
            HSSFSheet SheetOne = (HSSFSheet)workbookWrite.GetSheet("Sheet1");
            HSSFSheet SheetTwo = (HSSFSheet)workbookWrite.GetSheet("Sheet2");

            if (filenameRead.Contains(".xls"))
            {
                workbookRead = new HSSFWorkbook(filestreamRead);
            }
            else if (filenameRead.Contains(".xlsx"))
            {
                workbookRead = new XSSFWorkbook(filestreamRead);
            }
            
            // 创建一个符合条件的列表：包含（1M001，1M002等）
            List<string> CellText = new List<string>();
            for (int i = 1; i <= 4; i++)
            {
                for(int j = 1; j <= 3; j++)
                {
                    string codeM = $"{i}{Section.M}00{j}";
                    //string codeF = i + Section.F + "00" + j;
                    string codeF = $"{i}{Section.F}00{j}";
                    CellText.Add(codeM);
                    CellText.Add(codeF);
                }
            }

            List<string> CellMulText = new List<string>();
            CellMulText.Add("Total Consumed");
            CellMulText.Add("Avg / Day");

            ISheet sheet1 = workbookRead.GetSheetAt(0);
            IRow row;
            HSSFRow Row, Row1,Row2;
            HSSFCell[] sheetCell = new HSSFCell[13];
            HSSFCell[] SheetOneCell = new HSSFCell[13];
            HSSFCell[] SheetTwoCell = new HSSFCell[13];

            // 创建标题行
            SheetOne.CreateRow(0);
            Row1 = (HSSFRow)SheetOne.GetRow(0);
            for(int i = 0; i <= SheetOneCell.Length-1; i++)
            {
                SheetOneCell[i] = (HSSFCell)Row1.CreateCell(i);
            }
            SheetOneCell[0].SetCellValue("Animal");
            SheetOneCell[1].SetCellValue("Pretest(Rodent)");
            SheetOneCell[2].SetCellValue("Pretest(Rodent)");
            SheetOneCell[3].SetCellValue("D7");
            SheetOneCell[4].SetCellValue("D7");
            SheetOneCell[5].SetCellValue("D14");
            SheetOneCell[6].SetCellValue("D14");
            SheetOneCell[7].SetCellValue("D21");
            SheetOneCell[8].SetCellValue("D21");
            SheetOneCell[9].SetCellValue("D28");
            SheetOneCell[10].SetCellValue("D28");
            SheetOneCell[11].SetCellValue("D35");
            SheetOneCell[12].SetCellValue("D35");

            // 创建sheet2标题行
            SheetTwo.CreateRow(0);
            Row2 = (HSSFRow)SheetTwo.GetRow(0);
            for (int i = 0; i <= SheetTwoCell.Length - 1; i++)
            {
                SheetTwoCell[i] = (HSSFCell)Row2.CreateCell(i);
            }
            SheetTwoCell[0].SetCellValue("Animal");
            SheetTwoCell[1].SetCellValue("Pretest(Rodent)");
            SheetTwoCell[2].SetCellValue("Pretest(Rodent)");
            SheetTwoCell[3].SetCellValue("R7");
            SheetTwoCell[4].SetCellValue("R7");
            SheetTwoCell[5].SetCellValue("R14");
            SheetTwoCell[6].SetCellValue("R14");
            SheetTwoCell[7].SetCellValue("R21");
            SheetTwoCell[8].SetCellValue("R21");
            SheetTwoCell[9].SetCellValue("R28");
            SheetTwoCell[10].SetCellValue("R28");
            SheetTwoCell[11].SetCellValue("R35");
            SheetTwoCell[12].SetCellValue("R35");

            for (int i = 1; i <= sheet1.LastRowNum; i++)
            {
                row = sheet1.GetRow(i);
                SheetOne.CreateRow(i);
                Row = (HSSFRow)SheetOne.GetRow(i);
                if (row != null)
                {
                    for (int j = 0; j < CellText.Count; j++)
                    {
                        try
                        {
                            //if ((row.GetCell(0).ToString().Contains(CellText[j])) || (row.GetCell(0).ToString().Contains(CellMulText[0])) || (row.GetCell(0).ToString().Contains(CellMulText[1]))||(row.GetCell(0).ToString().Contains("Study Phase: Recovery Phase")))
                            if (row.GetCell(0).ToString().Contains(CellText[j]))
                            {
                                string cellValue = row.GetCell(0).ToString();
                                sheetCell[0] = (HSSFCell)Row.CreateCell(0);
                                sheetCell[0].SetCellValue(cellValue);
                                //Console.WriteLine(cellValue);
                            }
                        }
                        catch { }
                    }
                    try
                    {
                        if (row.GetCell(0).ToString().Contains(CellMulText[0]))
                        {
                            string cellValue = row.GetCell(0).ToString();
                            sheetCell[1] = (HSSFCell)Row.CreateCell(1);
                            sheetCell[1].SetCellValue(cellValue);
                            //Console.WriteLine(cellValue);
                        }
                    }
                    catch { }
                    try
                    {
                        if (row.GetCell(0).ToString().Contains(CellMulText[1]))
                        {
                            string cellValue = row.GetCell(0).ToString();
                            sheetCell[2] = (HSSFCell)Row.CreateCell(2);
                            sheetCell[2].SetCellValue(cellValue);
                            //Console.WriteLine(cellValue);
                        }
                    }
                    catch { }
                }    
            }

            Console.WriteLine("End");
            Console.ReadKey();
            FileStream fileStreamWrite = new FileStream(@"C:\Users\CHAOCHEN\Desktop\test01.xls", FileMode.Create);
            workbookWrite.Write(fileStreamWrite);
            fileStreamWrite.Close();
            filestreamRead.Close();
            workbookWrite.Close();
            workbookRead.Close();
        }
    }
}

