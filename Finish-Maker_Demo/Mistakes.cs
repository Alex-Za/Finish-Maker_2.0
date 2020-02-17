using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Finish_Maker_Demo
{
    class Mistakes
    {
        public Mistakes(FileReader reader, List<List<string>> pathes, bool skuFromPDCheck)
        {
            fileReader = reader;
            this.pathes = pathes;
            this.skuFromPDCheck = skuFromPDCheck;
        }

        private List<List<string>> pathes;
        private FileReader fileReader;

        private string criticalErrors;
        private string otherErrors;
        private bool skuFromPDCheck;
        
        public string CriticalErrors
        {
            get
            {
                if (criticalErrors == null)
                {
                    CheckForErrors();
                }
                return criticalErrors;
            }
        }
        public string OtherErrors
        {
            get
            {
                if (otherErrors == null)
                {
                    CheckForErrors();
                }
                return otherErrors;
            }
        }

        private void CheckForErrors()
        {
            for (int i = 0; i < 3; i++)
            {
                if (pathes[i] == null || pathes[i].Count == 0)
                {
                    criticalErrors += "Не выбраны все необходимые файлы" + Environment.NewLine;
                    return;
                }
            }


            for (int i = 0; i < 4; i += 3)
            {
                foreach (string s in pathes[i])
                {
                    if (Path.GetExtension(s) != ".csv")
                    {
                        criticalErrors += "Файл ексопрт линков имеет некорректное расширение, измените путь на формат .csv" + Environment.NewLine;
                        return;
                    }
                }
            }
            

            foreach (string s in pathes[1])
            {
                if (Path.GetExtension(s) != ".xlsx" && Path.GetExtension(s) != ".csv")
                {
                    criticalErrors += "Файл PD имеет некорректное расширение, измените путь на формат .xlsx или .csv" + Environment.NewLine;
                    return;
                }
            }

            foreach (string s in pathes[2])
            {
                if (Path.GetExtension(s) != ".xlsx" && Path.GetExtension(s) != ".csv")
                {
                    criticalErrors += "Файл ID имеет некорректное расширение, измените путь на формат .csv или .xlsx" + Environment.NewLine;
                    return;
                }
            }


            foreach (string path in pathes[1])
            {
                if (Path.GetExtension(path) == ".xlsx")
                {
                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, false))
                    {
                        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

                        if (workbookPart.Workbook.Sheets.Count() < 2)
                        {
                            criticalErrors += "В файле продукт даты нет необходимых листов" + Environment.NewLine;
                            return;
                        }

                        Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().ElementAt(0);
                        Worksheet worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;
                        SheetData sheetData = worksheet.Elements<SheetData>().First();

                        int headerColumnCount = sheetData.ElementAt(0).ChildElements.Count;

                        if (headerColumnCount != 11)
                        {
                            criticalErrors += "Некорректное количество колонок на первом листе продукт даты" + Environment.NewLine;
                            return;
                        }

                        Sheet sheet2 = workbookPart.Workbook.Descendants<Sheet>().ElementAt(1);
                        Worksheet worksheet2 = ((WorksheetPart)workbookPart.GetPartById(sheet2.Id)).Worksheet;
                        SheetData sheetData2 = worksheet2.Elements<SheetData>().First();

                        int headerColumnCount2 = sheetData2.ElementAt(0).ChildElements.Count;

                        if (headerColumnCount2 != 5)
                        {
                            criticalErrors += "Некорректное количество колонок на втором листе продукт даты" + Environment.NewLine;
                            return;
                        }

                        for (int i = 1; i < fileReader.PData.PDData1.Count; i++)
                        {
                            if (fileReader.PData.PDData1[i][0] == "")
                            {
                                criticalErrors += "Пустое значение на первом листе продукт даты в колонке Brand, строка " + i + Environment.NewLine;
                            }
                            if (fileReader.PData.PDData1[i][1] == "")
                            {
                                criticalErrors += "Пустое значение на первом листе продукт даты в колонке SKU, строка " + i + Environment.NewLine;
                            }
                            if (fileReader.PData.PDData1[i][5] == "" || fileReader.PData.PDData1[i][6] == "")
                            {
                                criticalErrors += "Пустое значение на первом листе продукт даты в колонках CategoryName or SubtypeName, строка " + i + Environment.NewLine;
                            }
                        }

                        if (criticalErrors != null)
                        {
                            return;
                        }
                    }
                }
                
            }

            HashSet<string> differentRegistrBrand = new HashSet<string>();

            for (int i = 1; i < fileReader.PData.PDData1.Count - 2; i++)
            {
                if (fileReader.PData.PDData1[i][0].ToLower() == fileReader.PData.PDData1[i + 1][0].ToLower() &&
                fileReader.PData.PDData1[i][0] != fileReader.PData.PDData1[i + 1][0])
                {
                    differentRegistrBrand.Add(fileReader.PData.PDData1[i][0]);
                }
            }

            if (differentRegistrBrand.Count != 0)
            {
                otherErrors += "Некоторые бренды содержат разный регистр букв, что может привести к ошибкам (отвалится MMY если один и тот же бренд записан с разным регистром букв в фитмент и продукт дате)" +
                    Environment.NewLine + "Бренды в которых есть такая ошибка: " + Environment.NewLine;

                foreach (string s in differentRegistrBrand)
                {
                    otherErrors += s + Environment.NewLine;
                }
            }

            if (skuFromPDCheck == true)
            {
                List<List<string>> exportLinks = fileReader.ExportLinks.ToList();
                int brandKayPosition = exportLinks[0].Count - 1;
                int pDataBrandKayPosition = fileReader.PData.PDData1[0].Count - 1;

                HashSet<string> allSKU = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (int i = 1; i < exportLinks.Count; i++)
                {
                    allSKU.Add(exportLinks[i][brandKayPosition]);
                }

                HashSet<string> missingSKU = new HashSet<string>();

                for (int i = 1; i < fileReader.PData.PDData1.Count - 1; i++)
                {
                    if (!allSKU.Contains(fileReader.PData.PDData1[i][pDataBrandKayPosition]))
                    {
                        missingSKU.Add(fileReader.PData.PDData1[i][1]);
                    }
                }

                if (missingSKU.Count != 0)
                {
                    otherErrors += "Согласно выбраному чекбоксу собрать финиш файл по SKU указаным в продукт дате." + Environment.NewLine + "Не все ску были залиты на сайт, ску которые есть в файле пд и нет в експорт линках:" + Environment.NewLine;

                    foreach (string s in missingSKU)
                    {
                        otherErrors += s + Environment.NewLine;
                    }
                }
            }

            //string[] columnNames = { "Product ID", "Brand", "SKU", "Product Name", "Child Title", "Images", "MMY", "Make", "Manufacturer ID", "Model", "Template", "Years", "linkwww" };

            //List<List<string>> exportLinks = fileReader.ExportLinks.ToList();

            //foreach (string column in columnNames)
            //{
            //    if (!exportLinks[0].Contains(column))
            //    {
            //        criticalErrors += "В файле експорт линков нет колонки - " + column + Environment.NewLine;
            //    }
            //}
        }

    }
}
