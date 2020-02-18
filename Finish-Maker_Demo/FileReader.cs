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
            NeededColumnNameInfo neededColumnNameInfo = null;
            var headerInitialize = false;

            HashSet<string> allSKUInPD = GetSKUFromPD();

            foreach (string filePath in path)
            {
                var dataFromFile = GetLineFromFile(filePath);

                foreach (var line in dataFromFile)
                {
                    if (!headerInitialize)
                    {
                        neededColumnNameInfo = GetNeededColumnNamesInfo(line);
                        headerInitialize = true;
                        var headers = neededColumnNameInfo.HeaderNames;
                        headers.Add("Brand+SKU");
                        yield return headers;
                        continue;
                    }

                    if (line[2] == "")
                        continue;

                    var dataLine = new List<string>();

                    for (int i = 0; i < neededColumnNameInfo.ColumnNumbers.Count; i++)
                    {
                        for (int x = 0; x < line.Length; x++)
                        {

                            if (neededColumnNameInfo.ColumnNumbers[i] == x)
                            {
                                dataLine.Add(line[x]);
                                break;
                            }
                        }
                    }

                    GetExportLinksInfo(dataLine);
                    if (skuFromPDCheck && !allSKUInPD.Contains(dataLine[dataLine.Count - 1]))
                        continue;

                    yield return dataLine;
                }
            }

        }
        
        private IEnumerable<string[]> GetLineFromFile(string path)
        {
            using (StreamReader reader = new StreamReader(path))
                while (!reader.EndOfStream)
                    yield return reader.ReadLine().Split('|');
        }

        private void ForParseIDcsv(string filePath, ProductID productID)
        {
            using (StreamReader reader = new StreamReader(filePath))
            {
                var firstLine = reader.ReadLine().Split('|');
                if (firstLine.Count() > 1)
                {
                    int idIndex = 0;
                    int makeIndex = 0;
                    int modelIndex = 0;
                    int yearsIndex = 0;
                    for (int i = 0; i < firstLine.Length; i++)
                    {
                        switch (firstLine[i])
                        {
                            case "Product ID":
                                idIndex = i;
                                break;
                            case "Make":
                                makeIndex = i;
                                break;
                            case "Model":
                                modelIndex = i;
                                break;
                            case "Years":
                                yearsIndex = i;
                                break;
                        }
                    }

                    while (!reader.EndOfStream)
                    {
                        if (productID.ProdIDMMY == null)
                        {
                            productID.ProdIDMMY = new HashSet<string>();
                        }

                        var line = reader.ReadLine().Split('|');
                        productID.ProdID.Add(line[idIndex]);
                        productID.ProdIDMMY.Add(line[idIndex] + '|' + line[makeIndex] + '|' + line[modelIndex] + '|' + line[yearsIndex]);
                    }
                }
                else
                {
                    while (!reader.EndOfStream)
                    {
                        var line = reader.ReadLine();
                        productID.ProdID.Add(line);
                    }
                }
            }
        }
        private void ForParseIDxlsx(string filePath, ProductID productID)
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
                int makeIndex = 0;
                int modelIndex = 0;
                int yearsIndex = 0;
                for (int i = 0; i < dataList.Count; i++)
                {
                    switch (dataList[i])
                    {
                        case "Product ID":
                            idIndex = i;
                            break;
                        case "Make":
                            makeIndex = i;
                            break;
                        case "Model":
                            modelIndex = i;
                            break;
                        case "Years":
                            yearsIndex = i;
                            break;
                    }
                }

                if (productID.ProdIDMMY == null)
                {
                    productID.ProdIDMMY = new HashSet<string>();
                }

                foreach (Row row in sheetData.Elements<Row>())
                {
                    Cell cell = row.Descendants<Cell>().ElementAt(idIndex - 1);
                    Cell cell2 = row.Descendants<Cell>().ElementAt(makeIndex - 1);
                    Cell cell3 = row.Descendants<Cell>().ElementAt(modelIndex - 1);
                    Cell cell4 = row.Descendants<Cell>().ElementAt(yearsIndex - 1);

                    productID.ProdID.Add(GetCellValue(spreadsheetDocument, cell));
                    productID.ProdIDMMY.Add(GetCellValue(spreadsheetDocument, cell) + '|' + GetCellValue(spreadsheetDocument, cell2) + '|' + GetCellValue(spreadsheetDocument, cell3) + '|' + GetCellValue(spreadsheetDocument, cell4));
                }
            }
        }
        private ProductID ParsIDs(List<string> path)
        {

            ProductID productID = new ProductID();
            productID.ProdIDMMY = null;
            productID.ProdID = new HashSet<string>();

            foreach (string filePath in path)
            {
                if (Path.GetExtension(filePath) == ".csv")
                {
                    ForParseIDcsv(filePath, productID);
                }
                else if (Path.GetExtension(filePath) == ".xlsx")
                {
                    ForParseIDxlsx(filePath, productID);
                }
            }

            return productID;
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

        private void ForParseProductDataxlsx(string filePath, ProductData productData)
        {
            List<List<string>> dataSheet1 = new List<List<string>>();
            List<List<string>> dataSheet2 = new List<List<string>>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                if (productData.PDData1 != null)
                {
                    dataSheet1 = ParseExcelSheet(spreadsheetDocument, 0, true);
                    AddKeyPDSheet1(dataSheet1);
                    dataSheet2 = ParseExcelSheet(spreadsheetDocument, 1, true);
                    AddKeyPDSheet2(dataSheet2);
                }
                else
                {
                    dataSheet1 = ParseExcelSheet(spreadsheetDocument, 0, false);
                    AddKeyPDSheet1(dataSheet1);
                    dataSheet2 = ParseExcelSheet(spreadsheetDocument, 1, false);
                    AddKeyPDSheet2(dataSheet2);
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
        private void ForParseProductDatacsv(string filePath, ProductData productData)
        {
            List<List<string>> dataSheet1 = new List<List<string>>();
            List<List<string>> dataSheet2 = new List<List<string>>();
            bool headerInitialize = false;
            if (productData.PDData1 != null)
            {
                headerInitialize = true;
            }

            using (StreamReader reader = new StreamReader(filePath))
            {
                while (!reader.EndOfStream)
                {
                    if (headerInitialize)
                    {
                        reader.ReadLine();
                        headerInitialize = false;
                        continue;
                    }
                    var line = reader.ReadLine().Split('|');
                    List<string> dataLine1 = new List<string>();
                    List<string> dataLine2 = new List<string>();
                    for (int i = 0; i < 11; i++)
                    {
                        dataLine1.Add(line[i]);
                    }

                    if (line[11] != "")
                    {
                        for (int i = 11; i < 16; i++)
                        {
                            dataLine2.Add(line[i]);
                        }
                        dataSheet2.Add(dataLine2);
                    }
                    dataSheet1.Add(dataLine1);
                }
            }

            AddKeyPDSheet1(dataSheet1);
            AddKeyPDSheet2(dataSheet2);

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
        private ProductData ParsProductData(List<string> path)
        {
            ProductData productData = new ProductData();

            foreach (string filePath in path)
            {
                if (Path.GetExtension(filePath) == ".xlsx")
                {
                    ForParseProductDataxlsx(filePath, productData);
                }
                else if (Path.GetExtension(filePath) == ".csv")
                {
                    ForParseProductDatacsv(filePath, productData);
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
                {
                    exportLinks = ParsExportLinks(pathes[0]);
                }
                return exportLinks;
            }
        }

        private ProductID id;
        public ProductID ID
        {
            get
            {
                if (id == null)
                {
                    id = ParsIDs(pathes[2]);
                }
                return id;
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

        private void GetExportLinksInfo(List<string> dataLine)
            => dataLine.Add(dataLine[1] + dataLine[2]);
        private NeededColumnNameInfo GetNeededColumnNamesInfo(string[] headers)
        {
            NeededColumnNameInfo columnNameInfo = new NeededColumnNameInfo();
            columnNameInfo.ColumnNumbers = new List<int>();
            columnNameInfo.HeaderNames = new List<string>();

            string[] neededColumNames = { "Product ID", "Brand", "SKU", "Product Name", "Child Title", "Images", "MMY", "Make", "Manufacturer ID", "Model", "Template", "Years", "linkwww" };

            for (int i = 0; i < neededColumNames.Length; i++)
            {
                for (int x = 0; x < headers.Length; x++)
                {
                    if (neededColumNames[i] == headers[x])
                    {
                        columnNameInfo.ColumnNumbers.Add(x);

                        columnNameInfo.HeaderNames.Add(headers[x]);
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
            public List<string> HeaderNames { get; set; }
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

        private void AddKeyPDSheet1(List<List<string>> sheetData1)
        {
            foreach (List<string> lineData in sheetData1)
            {
                //if (lineData[0] == "Standard®" && lineData[10] != "")
                //{
                //    lineData.Add(lineData[0].Remove(lineData[0].Length - 1) + " - " + lineData[10]);
                //}

                lineData.Add(lineData[0].Remove(lineData[0].Length - 1)); // добавление колонки бренд + серия

                lineData.Add(lineData[0] + lineData[1]); // добавление колонки бренд + ску
            }
        }
        private void AddKeyPDSheet2(List<List<string>> sheetData2)
        {
            foreach (List<string> lineData in sheetData2)
            {
                lineData.Add(lineData[0].Remove(lineData[0].Length - 1));
            }
        }
    }

    class ProductData
    {
        public List<List<string>> PDData1 { get; set; }
        public List<List<string>> PDData2 { get; set; }
    }
}
