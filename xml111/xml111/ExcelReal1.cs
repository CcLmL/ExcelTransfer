using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace xml111
{
    class ExcelReal1
    {
        static void Main(string[] args)
        {
            // 创建读取的对象
            IWorkbook workbookRead = null;
            //string filenameRead = @"C:\Users\CHAOCHEN\Desktop\shl.xls";
            Console.WriteLine("请输入文件路径：");
            string filenameRead = Console.ReadLine();
            //Console.WriteLine(filenameRead1);
            Console.WriteLine(filenameRead);
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

            // 创建一个符合条件的列表：包含（1M001，1M002等）
            List<string> CellText = new List<string>();
            for (int i = 1; i <= 5; i++)
            {
                for (int j = 1; j <= 100; j++)
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
            HSSFRow Row, Row1, Row2;
            HSSFCell[] sheetCell = new HSSFCell[13];
            HSSFCell[] SheetOneCell = new HSSFCell[13];
            HSSFCell[] SheetTwoCell = new HSSFCell[13];

            int DosingPhaseRow = rowNum(workbookRead, "Study Phase: Dosing Phase");
            int RecoveryPhaseRow = rowNum(workbookRead, "Study Phase: Recovery Phase");
            int PretestPhaseRow = rowNum(workbookRead, "Study Phase: Pretest( Rodent )");
            int rownum = 1;
            int rownum1 = 1;
            int rownum2 = 1;
            int rownum7 = 1;

            // 创建标题行
            SheetOne.CreateRow(0);
            Row1 = (HSSFRow)SheetOne.GetRow(0);
            for (int i = 0; i <= SheetOneCell.Length - 1; i++)
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

            //PretestPhase
            for (int i = PretestPhaseRow; i <= DosingPhaseRow; i++)
            {
                //row = sheet1.GetRow(i);
                //SheetOne.CreateRow(i);
                //Row = (HSSFRow)SheetOne.GetRow(i);
                if (GetRow(sheet1, i) != null)
                {
                    for (int j = 0; j < CellText.Count; j++)
                    {
                        try
                        {
                            if(GetRow(sheet1, i).GetCell(0).ToString().Contains(CellText[j]))
                            {
                                string cellValue = GetRow(sheet1, i).GetCell(0).ToString();
                                SheetOne.CreateRow(rownum);
                                Row = (HSSFRow)SheetOne.GetRow(rownum);
                                sheetCell[0] = (HSSFCell)Row.CreateCell(0);
                                sheetCell[0].SetCellValue(cellValue);
                                //int length = Encoding.UTF8.GetBytes(sheetCell[0].ToString()).Length;// 获取当前单元格宽度
                                //SheetOne.SetColumnWidth(0, length * 256);
                                rownum++;
                            }
                        }
                        catch { }
                    }
                    try
                    {
                        if (GetRow(sheet1, i).GetCell(0).ToString().Contains(CellMulText[0]))
                        {
                            string cellValue = GetRow(sheet1, i).GetCell(0).ToString();
                            Row = (HSSFRow)SheetOne.GetRow(rownum1);
                            sheetCell[1] = (HSSFCell)Row.CreateCell(1);
                            sheetCell[1].SetCellValue(cellValue);
                            int length = Encoding.UTF8.GetBytes(sheetCell[1].ToString()).Length;// 获取当前单元格宽度
                            SheetOne.SetColumnWidth(1, length * 256);
                            rownum1++;
                        }
                    }
                    catch { }
                    try
                    {
                        if (GetRow(sheet1, i).GetCell(0).ToString().Contains(CellMulText[1]))
                        {
                            string cellValue = GetRow(sheet1, i).GetCell(0).ToString();
                            Row = (HSSFRow)SheetOne.GetRow(rownum2);
                            sheetCell[2] = (HSSFCell)Row.CreateCell(2);
                            sheetCell[2].SetCellValue(cellValue);
                            int length = Encoding.UTF8.GetBytes(sheetCell[2].ToString()).Length;// 获取当前单元格宽度
                            SheetOne.SetColumnWidth(2, length * 256);
                            rownum2++;
                        }
                    }
                    catch { }
                }
            }

            //DosingPhaseRow
            for (int i = DosingPhaseRow+1; i < RecoveryPhaseRow; i++)
            {
                if (GetRow(sheet1, i) != null)
                {
                    for (int j = 0; j < CellText.Count; j++)
                    {
                        try
                        {
                            if (GetRow(sheet1, i).GetCell(0).ToString().Contains(CellText[j]))
                            {
                                Row = (HSSFRow)SheetOne.GetRow(rownum7);

                                string TotalD;
                                string AvgD;

                                TotalD = GetRow(sheet1, i + 3).GetCell(0).ToString();
                                AvgD = GetRow(sheet1, i + 3 + 1).GetCell(0).ToString();
                                if (TotalD.Contains(CellMulText[0]) && AvgD.Contains(CellMulText[1]))
                                {
                                    SheetOneCell[1] = (HSSFCell)Row.CreateCell(3);
                                    SheetOneCell[2] = (HSSFCell)Row.CreateCell(4);
                                    SheetOneCell[1].SetCellValue(TotalD);
                                    SheetOneCell[2].SetCellValue(AvgD);
                                    int length = Encoding.UTF8.GetBytes(SheetOneCell[1].ToString()).Length;// 获取当前单元格宽度
                                    SheetOne.SetColumnWidth(3, length * 256);
                                    int length1 = Encoding.UTF8.GetBytes(SheetOneCell[1].ToString()).Length;// 获取当前单元格宽度
                                    SheetOne.SetColumnWidth(4, length * 256);
                                }
                                else
                                {
                                    SheetOneCell[1] = (HSSFCell)Row.CreateCell(3);
                                    SheetOneCell[2] = (HSSFCell)Row.CreateCell(4);
                                    SheetOneCell[1].SetCellValue("");
                                    SheetOneCell[2].SetCellValue("");
                                }
                                
                                for (int k = 0; k < 4; k++)
                                {
                                    TotalD = GetRow(sheet1, i + 3 + 4 * (k + 1)).GetCell(0).ToString();
                                    AvgD = GetRow(sheet1, i + 3 + 1 + 4 * (k + 1)).GetCell(0).ToString();
                                    if(TotalD.Contains(CellMulText[0]) &&AvgD.Contains(CellMulText[1]))
                                    {
                                        SheetOneCell[k] = (HSSFCell)Row.CreateCell(2 * k + 5);
                                        SheetOneCell[k + 1] = (HSSFCell)Row.CreateCell(2 * k + 6);
                                        SheetOneCell[k].SetCellValue(TotalD);
                                        SheetOneCell[k + 1].SetCellValue(AvgD);
                                        int length2 = Encoding.UTF8.GetBytes(SheetOneCell[1].ToString()).Length;// 获取当前单元格宽度
                                        SheetOne.SetColumnWidth(2 * k + 5, length2 * 256);
                                        int length3 = Encoding.UTF8.GetBytes(SheetOneCell[1].ToString()).Length;// 获取当前单元格宽度
                                        SheetOne.SetColumnWidth(2 * k + 6, length2 * 256);
                                    }
                                    else
                                    {
                                        SheetOneCell[k] = (HSSFCell)Row.CreateCell(2 * k + 5);
                                        SheetOneCell[k + 1] = (HSSFCell)Row.CreateCell(2 * k + 6);
                                        SheetOneCell[k].SetCellValue("");
                                        SheetOneCell[k + 1].SetCellValue("");
                                    }
                                }
                                rownum7++;
                            }
                        }
                        catch { }
                    }
                }
            }

            // 创建标题行
            SheetOne.CreateRow(rownum);
            Row2 = (HSSFRow)SheetOne.GetRow(rownum);
            for (int i = 0; i <= SheetTwoCell.Length - 1; i++)
            {
                SheetTwoCell[i] = (HSSFCell)Row2.CreateCell(i);
            }
            SheetTwoCell[0].SetCellValue("Animal");
            SheetTwoCell[1].SetCellValue("R7");
            SheetTwoCell[2].SetCellValue("R7");
            SheetTwoCell[3].SetCellValue("R14");
            SheetTwoCell[4].SetCellValue("R14");
            SheetTwoCell[5].SetCellValue("R21");
            SheetTwoCell[6].SetCellValue("R21");
            SheetTwoCell[7].SetCellValue("R28");
            SheetTwoCell[8].SetCellValue("R28");

            //RecoveryPhase
            for (int i = RecoveryPhaseRow + 1 ; i <= sheet1.LastRowNum; i++)
            {
                //row = sheet1.GetRow(i);
                //rowAnimal = sheet1.GetRow(i - 3);
                //Row = (HSSFRow)SheetOne.GetRow(i);
                if (GetRow(sheet1, i) != null)
                {
                    for (int j = 0; j < CellText.Count; j++)
                    {
                        try
                        {
                            if (GetRow(sheet1, i).GetCell(0).ToString().Contains(CellText[j]))
                            {
                                string cellValue = GetRow(sheet1, i).GetCell(0).ToString();
                                SheetOne.CreateRow(rownum + 1);
                                Row = (HSSFRow)SheetOne.GetRow(rownum + 1);
                                SheetOneCell[0] = (HSSFCell)Row.CreateCell(0);
                                SheetOneCell[0].SetCellValue(cellValue);

                                string TotalR;
                                string AvgR;

                                TotalR = GetRow(sheet1, i + 3).GetCell(0).ToString();
                                AvgR = GetRow(sheet1, i + 3 + 1).GetCell(0).ToString();

                                if (TotalR.Contains(CellMulText[0]) && AvgR.Contains(CellMulText[1]))
                                {
                                    SheetOneCell[1] = (HSSFCell)Row.CreateCell(1);
                                    SheetOneCell[2] = (HSSFCell)Row.CreateCell(2);
                                    SheetOneCell[1].SetCellValue(TotalR);
                                    SheetOneCell[2].SetCellValue(AvgR);
                                }
                                else
                                {
                                    SheetOneCell[1] = (HSSFCell)Row.CreateCell(1);
                                    SheetOneCell[2] = (HSSFCell)Row.CreateCell(2);
                                    SheetOneCell[1].SetCellValue("");
                                    SheetOneCell[2].SetCellValue("");
                                }

                                for (int k = 0; k < 3; k++)
                                {
                                    TotalR = GetRow(sheet1, i + 3 + 4 * (k + 1)).GetCell(0).ToString();
                                    AvgR = GetRow(sheet1, i + 3 + 1 + 4 * (k + 1)).GetCell(0).ToString();
                                    if(TotalR.Contains(CellMulText[0]) && AvgR.Contains(CellMulText[1]))
                                    {
                                        SheetOneCell[k] = (HSSFCell)Row.CreateCell(2 * k + 3);
                                        SheetOneCell[k + 1] = (HSSFCell)Row.CreateCell(2 * k + 4);
                                        SheetOneCell[k].SetCellValue(TotalR);
                                        SheetOneCell[k + 1].SetCellValue(AvgR);
                                    }
                                    else
                                    {
                                        SheetOneCell[k] = (HSSFCell)Row.CreateCell(2 * k + 3);
                                        SheetOneCell[k + 1] = (HSSFCell)Row.CreateCell(2 * k + 4);
                                        SheetOneCell[k].SetCellValue("");
                                        SheetOneCell[k + 1].SetCellValue("");
                                    }
                                }
                                rownum++;
                            }
                        }
                        catch { }
                    }
                }
            }

            Console.WriteLine("已经在桌面成功生成！按任意键以结束。");
            Console.ReadKey();
            FileStream fileStreamWrite = new FileStream(@"C:\Users\CHAOCHEN\Desktop\test.xls", FileMode.Create);
            workbookWrite.Write(fileStreamWrite);
            fileStreamWrite.Close();
            filestreamRead.Close();
            workbookWrite.Close();
            workbookRead.Close();
        }
        static int rowNum(IWorkbook workbookRead, string Text)
        {
            int row = 0;
            IRow row1;
            ISheet sheet1 = workbookRead.GetSheetAt(0);
            for(int i =1; i < sheet1.LastRowNum; i++)
            {
                row1 = sheet1.GetRow(i);
                try
                {
                    if (row1.GetCell(0).ToString().Contains(Text))
                    {
                        row = i;
                    }
                }
                catch { }
            }
            return row;
        }

        static IRow GetRow(ISheet sheet, int i)
        {
            IRow Row;
            Row = sheet.GetRow(i);
            return Row;
        }
    }
    enum Section
    {
        M,
        F
    }
}


