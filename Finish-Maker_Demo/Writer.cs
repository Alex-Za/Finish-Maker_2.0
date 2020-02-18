using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace Finish_Maker_Demo
{
    class Writer
    {
        private Processing processing;
        public Writer(Processing process, Action<int> ChangeProgress, string saveFilePath)
        {
            processing = process;
            this.changeProgress = ChangeProgress;
            this.saveFilePath = saveFilePath;
        }

        Action<int> changeProgress;
        private string saveFilePath;

        public void Write()
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook;
            Excel.Worksheet worksheet;
            Excel.Sheets sheet;
            app.Visible = false;
            app.DisplayAlerts = false;
            workbook = app.Workbooks.Add(Type.Missing);
            sheet = workbook.Sheets;

            //работа с первым листом книги
            worksheet = workbook.Worksheets[1];
            int lastFilledRow = 8;

            worksheet.Cells[1, 1].EntireColumn.ColumnWidth = 6;

            changeProgress(20);

            HeaderWrite(worksheet, lastFilledRow);

            changeProgress(25);

            JobberAppWrite(worksheet, lastFilledRow);

            changeProgress(30);

            lastFilledRow = lastFilledRow + processing.JobberApp.GetLength(0) + 1;

            MainWrite(worksheet, lastFilledRow);

            changeProgress(40);

            lastFilledRow = lastFilledRow + processing.MainData.GetLength(0) + 2;

            //изменить размер шрифта
            Excel.Range startCell = (Excel.Range)worksheet.Cells[1, 1];
            Excel.Range endCell = (Excel.Range)worksheet.Cells[lastFilledRow, 20];
            Excel.Range range = worksheet.Range[startCell, endCell];

            range.Font.Size = 14;

            //создание и запись листа с новыми товарами
            worksheet = (Excel.Worksheet)workbook.Sheets.Add(Type.Missing, sheet[1], Type.Missing, Type.Missing);

            WriteNewSKUList(worksheet);

            changeProgress(50);

            //создание листа новых серий
            worksheet = (Excel.Worksheet)workbook.Sheets.Add(Type.Missing, sheet[2], Type.Missing, Type.Missing);

            WriteNewSeriesList(worksheet);

            changeProgress(60);

            //создание и запись листа фитмент апдейта
            worksheet = (Excel.Worksheet)workbook.Sheets.Add(Type.Missing, sheet[3], Type.Missing, Type.Missing);

            WriteFitmUpdateList(worksheet);

            changeProgress(70);

            //создание и запись листа проблематик ску
            worksheet = (Excel.Worksheet)workbook.Sheets.Add(Type.Missing, sheet[4], Type.Missing, Type.Missing);

            WriteProblematicList(worksheet);

            changeProgress(80);

            //создание листа пендинг
            worksheet = (Excel.Worksheet)workbook.Sheets.Add(Type.Missing, sheet[5], Type.Missing, Type.Missing);
            worksheet.Name = "Pending Issues";

            //создание листа чаилд дубликатов
            worksheet = (Excel.Worksheet)workbook.Sheets.Add(Type.Missing, sheet[6], Type.Missing, Type.Missing);

            WriteChildDupList(worksheet);

            //финиш
            worksheet = workbook.Worksheets[1];
            worksheet.Activate();

            DateTime dateTime = DateTime.UtcNow.Date;
            string currentDate = dateTime.ToString("MM-dd-yy");
            string pathDir = saveFilePath;

            string categoryVal = CheckForSlash(processing.CategoryValue[0]);
            string userNameVal = CheckForSlash(processing.userName);
            string currentFileName = pathDir + "\\" + categoryVal + "_" + userNameVal + "_Updated-" + currentDate;
            int availableFileNumber = 1;
            while (File.Exists(currentFileName + ".xlsx"))
            {
                currentFileName = pathDir + "\\" + categoryVal + "_" + userNameVal + "_Updated-" + currentDate + "(" + availableFileNumber + ")";
                availableFileNumber++;
            }

            workbook.SaveAs(currentFileName + ".xlsx");
            workbook.Close();
            app.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            app = null;
            workbook = null;
            worksheet = null;

            //создание цсв файла mfr+sku
            string mfrCount = (processing.MfrPlusSKU.Count - 1).ToString();
            string mfrPlusSKUPath = pathDir + "\\mpn_" + mfrCount + ".csv";
            int availableMPnFIleNumber = 1;
            while (File.Exists(mfrPlusSKUPath))
            {
                mfrPlusSKUPath = pathDir + "\\mpn_" + mfrCount + "(" + availableMPnFIleNumber + ")" + ".csv";
                availableMPnFIleNumber++;
            }

            using (StreamWriter file = new StreamWriter(mfrPlusSKUPath))
            {
                foreach (string s in processing.MfrPlusSKU)
                {
                    file.WriteLine(s);
                }
            }

            CreateWbForTask(saveFilePath);

            changeProgress(90);

        }
        private string CheckForSlash(string value)
        {
            if (value.Contains("/"))
            {
                value = value.Replace("/", "");
                return value;
            }
            return value;
        }

        private void HeaderWrite(Excel.Worksheet worksheet, int lastFilledRow)
        {
            worksheet.Name = "General";
            string[,] header = processing.Header;
            Excel.Range startCell = (Excel.Range)worksheet.Cells[2, 2];
            Excel.Range endCell = (Excel.Range)worksheet.Cells[lastFilledRow, 12];
            Excel.Range range = worksheet.Range[startCell, endCell];
            range.Value = header;

            //кастомизация шапки
            for (int i = 0; i < 8; i++)
            {
                worksheet.Cells[i + 1, 2].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            }
            for (int i = 0; i < 6; i++)
            {
                worksheet.Cells[i + 3, 9].Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            }
            worksheet.Cells[2, 2].Font.Bold = true;
            endCell = (Excel.Range)worksheet.Cells[2, 20];
            range = worksheet.Range[startCell, endCell];
            range.Interior.Color = System.Drawing.Color.FromArgb(235, 241, 222);
            range.Merge();
            worksheet.Cells[3, 4].Font.Bold = true;
            worksheet.Cells[3, 9].Font.Bold = true;
            worksheet.Cells[3, 12].Font.Bold = true;
            worksheet.Range[worksheet.Cells[3, 4], worksheet.Cells[3, 8]].Interior.Color = System.Drawing.Color.Red;
            worksheet.Range[worksheet.Cells[3, 12], worksheet.Cells[3, 20]].Interior.Color = System.Drawing.Color.Red;
            worksheet.Range[worksheet.Cells[4, 4], worksheet.Cells[8, 4]].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            worksheet.Cells[1, 3].EntireColumn.ColumnWidth = 24;
            worksheet.Cells[1, 4].EntireColumn.ColumnWidth = 34;

            for (int i = 3; i < 9; i++)
            {
                worksheet.Range[worksheet.Cells[i, 2], worksheet.Cells[i, 3]].Merge();
            }

            for (int i = 3; i < 9; i++)
            {
                worksheet.Range[worksheet.Cells[i, 4], worksheet.Cells[i, 8]].Merge();
            }
        }

        private void JobberAppWrite(Excel.Worksheet worksheet, int lastFilledRow)
        {
            string[,] joberApp = processing.JobberApp;
            Excel.Range startCell = (Excel.Range)worksheet.Cells[lastFilledRow + 1, 2];
            Excel.Range endCell = (Excel.Range)worksheet.Cells[lastFilledRow + joberApp.GetLength(0), 20];
            Excel.Range range = worksheet.Range[startCell, endCell];
            range.Value = joberApp;

            //кастомизация джобер/апп
            worksheet.Cells[9, 2].Font.Bold = true;
            worksheet.Cells[9, 4].Font.Bold = true;
            worksheet.Cells[9, 10].Font.Bold = true;
            worksheet.Cells[9, 15].Font.Bold = true;
            worksheet.Range[worksheet.Cells[9, 2], worksheet.Cells[9, 20]].Interior.Color = System.Drawing.Color.FromArgb(235, 241, 222);
            //worksheet.Cells[9, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            worksheet.Cells[9, 4].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            worksheet.Cells[9, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            worksheet.Cells[9, 15].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

            for (int i = 9; i < joberApp.GetLength(0) + 9; i++)
            {
                worksheet.Range[worksheet.Cells[i, 2], worksheet.Cells[i, 3]].Merge();
                worksheet.Range[worksheet.Cells[i, 4], worksheet.Cells[i, 9]].Merge();
                worksheet.Range[worksheet.Cells[i, 10], worksheet.Cells[i, 14]].Merge();
                worksheet.Range[worksheet.Cells[i, 15], worksheet.Cells[i, 20]].Merge();
                worksheet.Cells[i, 1].EntireRow.RowHeight = 18;
            }
        }

        private void MainWrite(Excel.Worksheet worksheet, int lastFilledRow)
        {
            string[,] mainData = processing.MainData;

            //форматирование
            int[,] numbersMainData1 = new int[mainData.GetLength(0) - 2, 7];
            int[,] numbersMainData2 = new int[mainData.GetLength(0) - 2, 2];

            for (int i = 0; i < mainData.GetLength(0) - 2; i++)
            {
                int number;

                for (int y = 0; y < 7; y++)
                {
                    if (Int32.TryParse(mainData[i + 2, y + 3], out number))
                    {
                        numbersMainData1[i, y] = number;
                    }
                    else
                    {
                        numbersMainData1[i, y] = 0;
                    }
                }

                for (int q = 0; q < 2; q++)
                {
                    if (Int32.TryParse(mainData[i + 2, q + 11], out number))
                    {
                        numbersMainData2[i, q] = number;
                    }
                    else
                    {
                        numbersMainData2[i, q] = 0;
                    }
                }
            }

            //запись
            Excel.Range startCell = (Excel.Range)worksheet.Cells[lastFilledRow + 1, 2];
            int mainDataRowCount = mainData.GetLength(0);
            Excel.Range endCell = (Excel.Range)worksheet.Cells[lastFilledRow + mainDataRowCount, 20];
            Excel.Range range = worksheet.Range[startCell, endCell];

            range.Value = mainData;

            startCell = (Excel.Range)worksheet.Cells[lastFilledRow + 3, 5];
            endCell = (Excel.Range)worksheet.Cells[lastFilledRow + mainDataRowCount, 11];
            range = worksheet.Range[startCell, endCell];

            range.Value = numbersMainData1;

            //выравнивание по центру
            for (int i = 1; i < mainDataRowCount - 1; i++)
            {
                for (int x = 1; x < 8; x++)
                {
                    range.Cells[i, x].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
            }

            startCell = (Excel.Range)worksheet.Cells[lastFilledRow + 3, 13];
            endCell = (Excel.Range)worksheet.Cells[lastFilledRow + mainDataRowCount, 14];
            range = worksheet.Range[startCell, endCell];

            range.Value = numbersMainData2;

            //выравнивание по центру
            for (int i = 1; i < mainDataRowCount - 1; i++)
            {
                for (int x = 1; x < 3; x++)
                {
                    range.Cells[i, x].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                }
            }

            //кастомизация
            startCell = (Excel.Range)worksheet.Cells[lastFilledRow + 1, 2];
            endCell = (Excel.Range)worksheet.Cells[lastFilledRow + mainDataRowCount, 20];
            range = worksheet.Range[startCell, endCell];

            startCell = (Excel.Range)range.Cells[1, 1];
            endCell = (Excel.Range)range.Cells[1, 19];
            startCell.Font.Bold = true;
            worksheet.Range[startCell, endCell].Merge();
            worksheet.Range[startCell, endCell].Interior.Color = System.Drawing.Color.FromArgb(235, 241, 222);
            for (int i = 1; i < 19; i++)
            {
                range.Cells[2, i].Font.Bold = true;
                range.Cells[2, i].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                range.Cells[2, i].WrapText = true;
            }
            range.Cells[2, 1].EntireRow.RowHeight = 55;
            startCell = (Excel.Range)range.Cells[2, 1];
            endCell = (Excel.Range)range.Cells[mainDataRowCount, 1];
            worksheet.Range[startCell, endCell].Interior.Color = System.Drawing.Color.FromArgb(235, 241, 222);

            startCell = (Excel.Range)range.Cells[2, 12];
            endCell = (Excel.Range)range.Cells[mainDataRowCount, 12];
            worksheet.Range[startCell, endCell].Interior.Color = System.Drawing.Color.FromArgb(235, 241, 222);

            startCell = (Excel.Range)range.Cells[2, 4];
            endCell = (Excel.Range)range.Cells[mainDataRowCount, 11];
            worksheet.Range[startCell, endCell].Interior.Color = System.Drawing.Color.FromArgb(252, 213, 180);

            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

            startCell = (Excel.Range)range.Cells[3, 1]; ;
            endCell = (Excel.Range)range.Cells[mainDataRowCount, 1];
            worksheet.Range[startCell, endCell].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


            //создание инвентари

            //lastFilledRow = lastFilledRow + mainDataRowCount;
            //startCell = (Excel.Range)worksheet.Cells[lastFilledRow + 1, 2];
            //endCell = (Excel.Range)worksheet.Cells[lastFilledRow + 2, 7];
            //range = worksheet.Range[startCell, endCell];

            //range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            //range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            //range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            //range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;

            //range.Interior.Color = Color.FromArgb(235, 241, 222);

            //range.Cells[2, 1].Value = "Inventory statistics";
            //range.Cells[2, 1].Font.Bold = true;
            //range.Merge();
            //range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            for (int i = 5; i < 14; i++)
            {
                worksheet.Cells[1, i].EntireColumn.ColumnWidth = 14;
            }

        }

        private void WriteNewSKUList(Excel.Worksheet worksheet)
        {
            worksheet.Name = "Added_SKUs";

            Excel.Range startCell = (Excel.Range)worksheet.Cells[1, 1];
            Excel.Range endCell = (Excel.Range)worksheet.Cells[processing.NewSKUList.GetLength(0), processing.NewSKUList.GetLength(1)];
            Excel.Range range = worksheet.Range[startCell, endCell];

            string[,] newskuRange = new string[processing.NewSKUList.GetLength(0), processing.NewSKUList.GetLength(1)];
            for (int i = 0; i < newskuRange.GetLength(0); i++)
            {
                for (int x = 0; x < newskuRange.GetLength(1); x++)
                {
                    newskuRange[i, x] = processing.NewSKUList[i, x];
                }
            }

            range.Value = newskuRange;
            for (int i = 1; i < newskuRange.GetLength(1) + 1; i++)
            {
                range.Cells[1, i].Font.Bold = true;
                range.Cells[1, i].EntireColumn.ColumnWidth = 12;
            }
        }

        private void WriteNewSeriesList(Excel.Worksheet worksheet)
        {
            worksheet.Name = "New_Series";
            worksheet.Cells[1, 1].Value = "Brand";
            worksheet.Cells[1, 1].Font.Bold = true;
            worksheet.Cells[1, 2].Value = "SuperProduct Product Name";
            worksheet.Cells[1, 2].Font.Bold = true;
            worksheet.Cells[1, 3].Value = "Link";
            worksheet.Cells[1, 3].Font.Bold = true;
            for (int i = 1; i < 3; i++)
            {
                worksheet.Cells[1, i].EntireColumn.ColumnWidth = 12;
            }
        }

        private void WriteFitmUpdateList(Excel.Worksheet worksheet)
        {
            worksheet.Name = "Fitment Updates";

            Excel.Range startCell = (Excel.Range)worksheet.Cells[1, 1];
            Excel.Range endCell = (Excel.Range)worksheet.Cells[processing.FitmentUpdateList.GetLength(0), processing.FitmentUpdateList.GetLength(1)];
            Excel.Range range = worksheet.Range[startCell, endCell];

            string[,] fitmentRange = new string[processing.FitmentUpdateList.GetLength(0), processing.FitmentUpdateList.GetLength(1)];
            for (int i = 0; i < fitmentRange.GetLength(0); i++)
            {
                for (int x = 0; x < fitmentRange.GetLength(1); x++)
                {
                    fitmentRange[i, x] = processing.FitmentUpdateList[i, x];
                }
            }

            range.Value = fitmentRange;
            for (int i = 1; i < processing.FitmentUpdateList.GetLength(1) + 1; i++)
            {
                range.Cells[1, i].Font.Bold = true;
                range.Cells[1, i].EntireColumn.ColumnWidth = 12;
            }
        }

        private void WriteProblematicList(Excel.Worksheet worksheet)
        {
            worksheet.Name = "Problematic SKU's (reason)";

            Excel.Range startCell = (Excel.Range)worksheet.Cells[1, 1];
            Excel.Range endCell = (Excel.Range)worksheet.Cells[processing.ProblematicSKU.GetLength(0), processing.ProblematicSKU.GetLength(1)];
            Excel.Range range = worksheet.Range[startCell, endCell];

            string[,] problematicRange = new string[processing.ProblematicSKU.GetLength(0), processing.ProblematicSKU.GetLength(1)];
            for (int i = 0; i < problematicRange.GetLength(0); i++)
            {
                for (int x = 0; x < problematicRange.GetLength(1); x++)
                {
                    problematicRange[i, x] = processing.ProblematicSKU[i, x];
                }
            }

            range.Value = problematicRange;
            for (int i = 1; i < processing.ProblematicSKU.GetLength(1) + 1; i++)
            {
                range.Cells[1, i].Font.Bold = true;
                range.Cells[1, i].EntireColumn.ColumnWidth = 12;
            }
        }

        private void WriteChildDupList(Excel.Worksheet worksheet)
        {
            worksheet.Name = "Child titles duplication";

            if (processing.ChildTitleDuplicates == null)
            {
                return;
            }

            Excel.Range startCell = (Excel.Range)worksheet.Cells[1, 1];
            Excel.Range endCell = (Excel.Range)worksheet.Cells[processing.ChildTitleDuplicates.GetLength(0), processing.ChildTitleDuplicates.GetLength(1)];
            Excel.Range range = worksheet.Range[startCell, endCell];

            range.Value = processing.ChildTitleDuplicates;
            for (int i = 1; i < processing.ChildTitleDuplicates.GetLength(1) + 1; i++)
            {
                range.Cells[1, i].Font.Bold = true;
                range.Cells[1, i].EntireColumn.ColumnWidth = 12;
            }
        }

        private void CreateWbForTask(string saveFilePath)
        {
            string pathDir = saveFilePath;
            string currentPath = pathDir + "\\" + "Brands.xlsx";
            int nextPath = 1;

            while (File.Exists(currentPath))
            {
                currentPath = pathDir + "\\" + "Brands(" + nextPath + ").xlsx";
                nextPath++;
            }

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(currentPath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                workbookPart.Workbook.Sheets = new Sheets();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>();

                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                Sheet sheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = sheetId,
                    Name = "mySheet"
                };
                sheets.Append(sheet);

                int brandsCount = processing.MainData.GetLength(0) - 2;

                string brands = "";

                for (int i = 2; i < processing.MainData.GetLength(0); i++)
                {
                    if (i == brandsCount + 1)
                    {
                        brands += processing.MainData[i, 1];
                        break;
                    }
                    brands += processing.MainData[i, 1] + ", ";
                }


                Row newRow = new Row();
                Cell cell = new Cell();
                cell.DataType = CellValues.String;
                cell.CellValue = new CellValue("Brands Count");
                newRow.AppendChild(cell);
                Cell cell2 = new Cell();
                cell2.DataType = CellValues.String;
                cell2.CellValue = new CellValue("Brands");
                newRow.AppendChild(cell2);
                sheetData.AppendChild(newRow);

                Row newRow2 = new Row();
                Cell cell3 = new Cell();
                cell3.DataType = CellValues.String;
                cell3.CellValue = new CellValue(brandsCount.ToString());
                newRow2.AppendChild(cell3);
                Cell cell4 = new Cell();
                cell4.DataType = CellValues.String;
                cell4.CellValue = new CellValue(brands);
                newRow2.AppendChild(cell4);
                sheetData.AppendChild(newRow2);

            }


            //Excel.Application app = new Excel.Application();
            //Excel.Workbook workbook;
            //Excel.Worksheet worksheet;
            //Excel.Sheets sheet;
            //app.Visible = false;
            //app.DisplayAlerts = false;
            //workbook = app.Workbooks.Add(Type.Missing);
            //sheet = workbook.Sheets;

            //worksheet = workbook.Worksheets[1];

            //worksheet.Cells[1, 1].Value = "Brands Count";
            //worksheet.Cells[1, 2].Value = "Brands";

            //int brandsCount = processing.MainData.GetLength(0) - 2;

            //string brands = "";

            //for (int i = 2; i < processing.MainData.GetLength(0); i++)
            //{
            //    if (i == brandsCount + 1)
            //    {
            //        brands += processing.MainData[i, 1];
            //        break;
            //    }
            //    brands += processing.MainData[i, 1] + ", ";
            //}

            //worksheet.Cells[2, 1].Value = brandsCount;
            //worksheet.Cells[2, 2].Value = brands;

            //string pathDir = Directory.GetCurrentDirectory();
            //string currentPath = pathDir + "\\" + "Brands.xlsx";
            //int nextPath = 1;

            //while (File.Exists(currentPath))
            //{
            //    currentPath = pathDir + "\\" + "Brands(" + nextPath + ").xlsx";
            //    nextPath++;
            //}

            //workbook.SaveAs(currentPath);
            //workbook.Close();
            //app.Quit();
        }

        public void WriteExcelFile(string filePath, HashSet<string> data)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                workbookPart.Workbook.Sheets = new Sheets();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                SheetData sheetData = new SheetData();
                worksheetPart.Worksheet = new Worksheet(sheetData);

                Sheets sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>();

                uint sheetId = 1;
                if (sheets.Elements<Sheet>().Count() > 0)
                {
                    sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
                }

                Sheet sheet = new Sheet()
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = sheetId,
                    Name = "mySheet"
                };
                sheets.Append(sheet);


                foreach (string s in data)
                {
                    Row newRow = new Row();
                    Cell cell = new Cell();
                    cell.DataType = CellValues.String;
                    cell.CellValue = new CellValue(s);
                    newRow.AppendChild(cell);
                    sheetData.AppendChild(newRow);
                }

                //foreach (List<string> dataList in data)
                //{
                //    Row newRow = new Row();
                //    foreach (string s in dataList)
                //    {
                //        Cell cell = new Cell();
                //        cell.DataType = CellValues.String;
                //        cell.CellValue = new CellValue(s);
                //        newRow.AppendChild(cell);
                //    }
                //    sheetData.AppendChild(newRow);
                //}

            }

        }
    }
}
