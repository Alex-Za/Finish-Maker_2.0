using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Finish_Maker_Demo
{
    class FileReader
    {
        private List<List<string>> pathes;
        private bool skuFromPDCheck;
        public FileReader(List<List<string>> pathes, bool skuFromPDCheck)
        {
            this.pathes = pathes;
            this.skuFromPDCheck = skuFromPDCheck;
        }

        private IEnumerable<List<string>> ParsExportLinks(List<string> path)
        {
            HashSet<string> transData = new HashSet<string>();

            NeededColumnNameInfo neededColumnNameInfo = null;
            var headerInitialize = false;

            foreach (string filePath in path)
            {
                var dataFromFile = GetLineFromFile(filePath);

                foreach (var line in dataFromFile)
                {
                    if (!headerInitialize)
                    {
                        neededColumnNameInfo = GetNeededColumnNamesInfo(line);
                        headerInitialize = true;
                        transData.Add(neededColumnNameInfo.HeaderNames);

                    }

                    if (line[2] == "")
                        continue;
                    var dataLine = string.Empty;

                    for (int i = 0; i < neededColumnNameInfo.ColumnNumbers.Count; i++)
                    {
                        for (int x = 0; x < line.Length; x++)
                        {
                            if (neededColumnNameInfo.ColumnNumbers[i] == x)
                            {
                                dataLine = (dataLine != "") ? dataLine + '|' + line[x] : line[x];
                                break;
                            }
                        }
                    }

                    transData.Add(dataLine);
                }
            }

            if (skuFromPDCheck == true)
            {
                HashSet<string> allSKUInPD = GetSKUFromPD();
                return GetInPDExportLinksInfo(transData, allSKUInPD);
            }
            else
            {
                return GetExportLinksInfo(transData);
            }


        }

        private IEnumerable<string[]> GetLineFromFile(string path)
        {
            using (StreamReader reader = new StreamReader(path))
                while (!reader.EndOfStream)
                    yield return reader.ReadLine().Split('|');
        }

        private HashSet<string> ParsIDs(List<string> path)
        {

            HashSet<string> idList = new HashSet<string>();

            foreach (string filePath in path)
            {
                if (Path.GetExtension(filePath) == ".csv")
                {
                    using (StreamReader reader = new StreamReader(filePath))
                    {
                        var firstLine = reader.ReadLine().Split('|');
                        if (firstLine.Count() > 1)
                        {
                            int productIDIndex = 0;
                            foreach (string s in firstLine)
                            {
                                if (s == "Product ID")
                                {
                                    break;
                                }
                                productIDIndex++;
                            }

                            while (!reader.EndOfStream)
                            {
                                var line = reader.ReadLine().Split('|');
                                idList.Add(line[productIDIndex]);
                            }
                        }
                        else
                        {
                            while (!reader.EndOfStream)
                            {
                                var line = reader.ReadLine();

                                idList.Add(line);

                            }
                        }

                    }
                }
                else if (Path.GetExtension(filePath) == ".xlsx")
                {
                    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                    {
                        List<List<string>> dataSheet = new List<List<string>>();

                        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                        Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().ElementAt(0);
                        Worksheet worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;
                        SheetData sheetData = worksheet.Elements<SheetData>().First();

                        Row firstRow = sheetData.Elements<Row>().First();

                        List<string> dataList = new List<string>();

                        for (int i = 0; i < firstRow.Descendants<Cell>().Count(); i++)
                        {
                            Cell cell = firstRow.Descendants<Cell>().ElementAt(i);
                            int actualCellIndex = CellReferenceToIndex(cell);
                            dataList.Add(GetCellValue(spreadsheetDocument, cell));
                        }

                        int idIndex = 0;
                        foreach (string s in dataList)
                        {
                            if (s == "Product ID")
                            {
                                break;
                            }
                            idIndex++;
                        }

                        foreach (Row row in sheetData.Elements<Row>())
                        {
                            Cell cell = row.Descendants<Cell>().ElementAt(idIndex - 1);
                            int actualCellIndex = CellReferenceToIndex(cell);
                            idList.Add(GetCellValue(spreadsheetDocument, cell));
                        }

                    }
                }

            }

            return idList;
        }

        private IEnumerable<List<string>> ParsChildTitleDuplicates(List<string> path)
        {
            foreach (string filePath in path)
            {
                var dataFromFile = GetLineFromFile(filePath);

                foreach (var line in dataFromFile)
                {
                    yield return line.ToList();
                }
            }
            
        }

        private ProductData ParsProductData(List<string> path)
        {
            ProductData productData = new ProductData();

            foreach (string filePath in path)
            {
                List<List<string>> dataSheet1 = new List<List<string>>();
                List<List<string>> dataSheet2 = new List<List<string>>();

                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
                {
                    if (productData.PDData1 != null)
                    {
                        dataSheet1 = ParseExcelSheet(spreadsheetDocument, 0, true);
                        dataSheet1 = AddKeyPDSheet1(dataSheet1);
                        dataSheet2 = ParseExcelSheet(spreadsheetDocument, 1, true);
                        dataSheet2 = AddKeyPDSheet2(dataSheet2);
                    }
                    else
                    {
                        dataSheet1 = ParseExcelSheet(spreadsheetDocument, 0, false);
                        dataSheet1 = AddKeyPDSheet1(dataSheet1);
                        dataSheet2 = ParseExcelSheet(spreadsheetDocument, 1, false);
                        dataSheet2 = AddKeyPDSheet2(dataSheet2);
                    }
                    
                }

                if (productData.PDData1 != null)
                {
                    productData.PDData1 = productData.PDData1.Concat(dataSheet1).ToList();
                    productData.PDData2 = productData.PDData2.Concat(dataSheet2).ToList();
                }
                else
                {
                    productData.PDData1 = dataSheet1;
                    productData.PDData2 = dataSheet2;
                }
            }

            return productData;
        }

        private IEnumerable<List<string>> exportLinks;
        public IEnumerable<List<string>> ExportLinks
        {
            get
            {
                if (exportLinks == null)
                    exportLinks = ParsExportLinks(pathes[0]).ToList();

                return exportLinks;
            }
        }
        private HashSet<string> iD;
        public HashSet<string> ID
        {
            get
            {
                if (iD == null)
                {
                    iD = ParsIDs(pathes[2]);
                }
                return iD;
            }
        }
        private List<List<string>> chtTitleDuplicates;
        public List<List<string>> ChtTitleDuplicates
        {
            get
            {
                if (chtTitleDuplicates == null)
                {
                    if (pathes[3].Count < 1)
                    {
                        chtTitleDuplicates = null;
                        return chtTitleDuplicates;
                    }
                    else
                    {
                        chtTitleDuplicates = ParsChildTitleDuplicates(pathes[3]).ToList();
                    }
                }
                return chtTitleDuplicates;
            }
        }


        private ProductData pData;
        public ProductData PData
        {
            get
            {
                if (pData == null)
                    pData = ParsProductData(pathes[1]);

                return pData;
            }
        }

        private IEnumerable<List<string>> GetExportLinksInfo(HashSet<string> transData)
        {
            List<string> firstLine = transData.First().Split('|').ToList();
            firstLine.Add("Brand+SKU");
            yield return firstLine;

            foreach (string data in transData.Skip(1))
            {
                List<string> dataLine = data.Split('|').ToList();
                dataLine.Add(dataLine[1] + dataLine[2]);
                yield return dataLine;
            }

        }
        private IEnumerable<List<string>> GetInPDExportLinksInfo(HashSet<string> transData, HashSet<string> allSKUInPD)
        {

            string firstLineTransData = transData.First();
            List<string> firstLineTransDataList = firstLineTransData.Split('|').ToList();

            firstLineTransDataList.Add("Brand+SKU");
            yield return firstLineTransDataList;

            foreach (string data in transData)
            {
                List<string> dataLine = data.Split('|').ToList();
                dataLine.Add(dataLine[1] + dataLine[2]);

                if (!allSKUInPD.Contains(dataLine[dataLine.Count - 1]))
                    continue;

                yield return dataLine;
            }

        }
        private NeededColumnNameInfo GetNeededColumnNamesInfo(string[] headers)
        {
            NeededColumnNameInfo columnNameInfo = new NeededColumnNameInfo();
            columnNameInfo.ColumnNumbers = new List<int>();

            string[] neededColumNames = { "Product ID", "Brand", "SKU", "Product Name", "Child Title", "Images", "MMY", "Make", "Manufacturer ID", "Model", "Template", "Years", "linkwww" };

            for (int i = 0; i < neededColumNames.Length; i++)
            {
                for (int x = 0; x < headers.Length; x++)
                {
                    if (neededColumNames[i] == headers[x])
                    {
                        columnNameInfo.ColumnNumbers.Add(x);

                        columnNameInfo.HeaderNames = (!String.IsNullOrEmpty(columnNameInfo.HeaderNames)) ? columnNameInfo.HeaderNames + '|' + headers[x] : headers[x];
                        break;
                    }
                }

            }

            return columnNameInfo;

        }
        private HashSet<string> GetSKUFromPD()
        {
            HashSet<string> allSKUInPD = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 1; i < PData.PDData1.Count; i++)
            {
                allSKUInPD.Add(PData.PDData1[i][PData.PDData1[0].Count - 1]);
            }

            return allSKUInPD;
        }
        class NeededColumnNameInfo
        {
            public string HeaderNames { get; set; }
            public List<int> ColumnNumbers { get; set; }
        }

        private int CellReferenceToIndex(Cell cell)
        {
            int index = 0;
            string reference = cell.CellReference.ToString().ToUpper();
            foreach (char ch in reference)
            {
                if (Char.IsLetter(ch))
                {
                    int value = (int)ch - (int)'A';
                    index = (index == 0) ? value : ((index + 1) * 26) + value;
                }
                else
                {
                    return index;
                }
            }
            return index;
        }
        private string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart sstpart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            SharedStringTable sst = sstpart.SharedStringTable;

            if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
            {
                int ssid = int.Parse(cell.CellValue.Text);
                string str = sst.ChildElements[ssid].InnerText;
                return str;
            }
            else if (cell.CellValue != null)
            {
                return cell.CellValue.Text;
            }
            else
            {
                return "";
            }
        }
        private List<List<string>> ParseExcelSheet(SpreadsheetDocument spreadsheetDocument, int sheetIndex, bool headerInitialize)
        {
            List<List<string>> dataSheet = new List<List<string>>();

            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().ElementAt(sheetIndex);
            Worksheet worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;
            SheetData sheetData = worksheet.Elements<SheetData>().First();


            int headerColumnCount = sheetData.ElementAt(0).ChildElements.Count;

            foreach (Row row in sheetData.Elements<Row>())
            {
                if (headerInitialize == true)
                {
                    headerInitialize = false;
                    continue;
                }

                List<string> dataList = new List<string>();

                for (int i = 0; i < headerColumnCount; i++)
                {
                    dataList.Add(string.Empty);
                }

                for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                {
                    Cell cell = row.Descendants<Cell>().ElementAt(i);
                    int actualCellIndex = CellReferenceToIndex(cell);
                    dataList[actualCellIndex] = GetCellValue(spreadsheetDocument, cell);
                }

                dataSheet.Add(dataList);

            }
            return dataSheet;
        }

        private List<List<string>> AddKeyPDSheet1(List<List<string>> sheetData1)
        {
            List<List<string>> sheetDataWithKey = sheetData1;
            foreach (List<string> lineData in sheetDataWithKey)
            {
                //if (lineData[0] == "Standard®" && lineData[10] != "")
                //{
                //    lineData.Add(lineData[0].Remove(lineData[0].Length - 1) + " - " + lineData[10]);
                //}

                lineData.Add(lineData[0].Remove(lineData[0].Length - 1)); // добавление колонки бренд + серия

                lineData.Add(lineData[0] + lineData[1]); // добавление колонки бренд + ску
            }

            return sheetDataWithKey;
        }
        private List<List<string>> AddKeyPDSheet2(List<List<string>> sheetData2)
        {
            List<List<string>> sheetData2WithKey = sheetData2;
            foreach (List<string> lineData in sheetData2WithKey)
            {
                lineData.Add(lineData[0].Remove(lineData[0].Length - 1));
            }

            return sheetData2WithKey;
        }
    }

    class ProductData
    {
        public List<List<string>> PDData1 { get; set; }
        public List<List<string>> PDData2 { get; set; }
    }
}
