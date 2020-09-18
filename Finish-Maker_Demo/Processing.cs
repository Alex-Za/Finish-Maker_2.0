using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Finish_Maker_Demo
{
    class Processing
    {
        private FileReader fileReader;
        public string userName;
        private bool fitmentUpdateCheck;
        private IEnumerable<List<string>> exportLinks;
        private ConsoleMessage message;
        SortedSet<string> brandSeriesKey = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
        public Processing(FileReader reader, string userName, bool fitmentUpdateCheck, ConsoleMessage message)
        {
            fileReader = reader;
            this.userName = userName;
            this.fitmentUpdateCheck = fitmentUpdateCheck;
            this.message = message;
        }

        private void GenerateMainData()
        {
            exportLinks = fileReader.ExportLinks;
            message.MessageTriger("Генерация мейн даты...");

            var isFirst = true;

            int problematicSKUPosition = 0;
            int fitmUpdatePosition = 0;
            int newIDCheckPosition = 0;
            int brandSeriesPosition = 0;
            int brandKeyPosition = 0;
            int pDataBrandSeriesPosition = fileReader.PData.PDData1[0].Count - 2;
            int pDataBrandKeyPosition = fileReader.PData.PDData1[0].Count - 1;
            int pData2BrandKeyPosition = fileReader.PData.PDData2[0].Count - 1;
            SortedSet<string> brands = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> allSKU = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> idsPlusBrandSeries = new HashSet<string>();
            HashSet<string> newSKU = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> fitmentUpdateSKU = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            HashSet<string> brandKeyPlusSKuFitmentUpdate = new HashSet<string>();
            HashSet<string> skuPlusBrandsSeries = new HashSet<string>();
            HashSet<string> problematicBrandSeriesPlusSKU = new HashSet<string>();

            GetNewAndFitmentUpdateSKU(newSKU, fitmentUpdateSKU, pDataBrandKeyPosition);

            foreach (var row in exportLinks)
            {

                if (isFirst)
                {
                    problematicSKUPosition = row.Count - 1;
                    fitmUpdatePosition = row.Count;
                    newIDCheckPosition = row.Count + 1;
                    brandSeriesPosition = row.Count + 2;
                    brandKeyPosition = row.Count + 3;
                    isFirst = false;

                    AddCheckColumnsToExpLinksHeader(row, problematicSKUPosition, fitmUpdatePosition, newIDCheckPosition, brandSeriesPosition);
                    continue;
                }

                AddCheckColumnsToExpLinksEmpty(row, problematicSKUPosition, fitmUpdatePosition, newIDCheckPosition, brandSeriesPosition);
                AddBrands(row, brandSeriesPosition, brandSeriesKey, brands);
                AddNewSKUCount(row, brandKeyPosition, allSKU);
                AddNewIDsCount(row, brandSeriesPosition, newIDCheckPosition, idsPlusBrandSeries);
                AddFitmentUpdate(row, brandSeriesPosition, brandKeyPosition, newIDCheckPosition, fitmUpdatePosition, brandKeyPlusSKuFitmentUpdate, newSKU, fitmentUpdateSKU);
                AddTotalWebsiteSKU(row, brandSeriesPosition, skuPlusBrandsSeries);
                AddProblematicSKU(row, brandSeriesPosition, problematicSKUPosition, problematicBrandSeriesPlusSKU);
                AddNewProductInfo(row, brandKeyPosition, newSKU);
                AddFitmentDataInfo(row, fitmUpdatePosition);
                AddProblematicInfo(row, problematicSKUPosition);
                AddMfrPlusSKU(row, newSKU, brandKeyPosition);
            }
            GenerateJobberApp(brands);
            GeneratingMainDataTabels(brandSeriesKey, pDataBrandSeriesPosition, allSKU, pDataBrandKeyPosition, idsPlusBrandSeries, brandKeyPlusSKuFitmentUpdate, skuPlusBrandsSeries, problematicBrandSeriesPlusSKU);

            AddImgPDfVideo(pDataBrandSeriesPosition);
            AddRanksAndRAPs(pData2BrandKeyPosition);
            AddImprovedStructure(pDataBrandSeriesPosition);
            AddUpdateType(pDataBrandKeyPosition);
            AddHeaderToMainData();
        }
        private void AddBrandsInMainData(SortedSet<string> brandSeriesKey, string[,] mainData)
        {
            int positionForMainData = 2;

            foreach (string s in brandSeriesKey)
            {
                mainData[positionForMainData, 1] = s;
                positionForMainData++;
            }
        }
        private void AddNewSkuCountInMainData(int pDataBrandSeriesPosition, int pDataBrandKeyPosition, HashSet<string> allSKU, string[,] mainData)
        {
            for (int i = 2; i < mainData.GetLength(0); i++)
            {
                int newProductCount = 0;
                for (int y = 1; y < fileReader.PData.PDData1.Count; y++)
                {
                    if (mainData[i, 1].ToLower() == fileReader.PData.PDData1[y][pDataBrandSeriesPosition].ToLower() && fileReader.PData.PDData1[y][2].ToLower() == "new" && allSKU.Contains(fileReader.PData.PDData1[y][pDataBrandKeyPosition]))
                    {
                        newProductCount++;
                    }
                }
                mainData[i, 3] = newProductCount.ToString();
            }
        }
        private void AddNewProdIDCountInMainData(HashSet<string> idsPlusBrandSeries, string[,] mainData)
        {
            List<string[]> brandIDsForCheck = new List<string[]>();

            foreach (string s in idsPlusBrandSeries)
            {
                string[] brandIDsLine = s.Split('|');
                brandIDsForCheck.Add(brandIDsLine);
            }

            for (int i = 2; i < mainData.GetLength(0); i++)
            {
                int newIDsCount = 0;
                for (int x = 0; x < brandIDsForCheck.Count; x++)
                {
                    if (mainData[i, 1].ToLower() == brandIDsForCheck[x][1].ToLower())
                    {
                        newIDsCount++;
                    }
                }
                mainData[i, 4] = newIDsCount.ToString();
            }
        }
        private void AddNewFitmUpdateCountInMainData(HashSet<string> brandKeyPlusSKuFitmentUpdate, string[,] mainData)
        {
            List<string[]> fitmentUpdateForCheck = new List<string[]>();
            string[] fitmentUpdateLine;

            foreach (string s in brandKeyPlusSKuFitmentUpdate)
            {
                fitmentUpdateLine = s.Split('|');
                fitmentUpdateForCheck.Add(fitmentUpdateLine);
            }

            for (int i = 2; i < mainData.GetLength(0); i++)
            {
                int fitmentUpdateCount = 0;
                for (int x = 0; x < fitmentUpdateForCheck.Count; x++)
                {
                    if (mainData[i, 1].ToLower() == fitmentUpdateForCheck[x][0].ToLower())
                    {
                        fitmentUpdateCount++;
                    }
                }
                mainData[i, 5] = fitmentUpdateCount.ToString();
            }
        }
        private void AddTotalSkuCountInMainData(HashSet<string> skuPlusBrandsSeries, string[,] mainData)
        {
            for (int i = 2; i < mainData.GetLength(0); i++)
            {
                int totalSKuCount = 0;
                foreach (string s in skuPlusBrandsSeries)
                {
                    string[] dataLine = s.Split('|');
                    if (mainData[i, 1].ToLower() == dataLine[0].ToLower())
                    {
                        totalSKuCount++;
                    }
                }
                mainData[i, 11] = totalSKuCount.ToString();
            }
        }
        private void AddProblemSkuCountInMainData(HashSet<string> problematicBrandSeriesPlusSKU, string[,] mainData)
        {
            for (int i = 2; i < mainData.GetLength(0); i++)
            {
                int problematicCount = 0;
                bool misImg = false;

                foreach (string s in problematicBrandSeriesPlusSKU)
                {
                    string[] problematicLine = s.Split('|');
                    if (problematicLine[0].ToLower() == mainData[i, 1].ToLower())
                    {
                        problematicCount++;
                        if (problematicLine[2] == "2" || problematicLine[2] == "3")
                        {
                            misImg = true;
                        }
                    }
                }

                mainData[i, 12] = problematicCount.ToString();
                if (misImg == true)
                {
                    mainData[i, 14] = "Task for Designers is still in progress";
                }
            }
        }
        private void GeneratingMainDataTabels(SortedSet<string> brandSeriesKey, int pDataBrandSeriesPosition, HashSet<string> allSKU, int pDataBrandKeyPosition, HashSet<string> idsPlusBrandSeries, HashSet<string> brandKeyPlusSKuFitmentUpdate, HashSet<string> skuPlusBrandsSeries, HashSet<string> problematicBrandSeriesPlusSKU)
        {
            mainData = new string[brandSeriesKey.Count + 2, 19];

            AddBrandsInMainData(brandSeriesKey, mainData);

            AddNewSkuCountInMainData(pDataBrandSeriesPosition, pDataBrandKeyPosition, allSKU, mainData);

            AddNewProdIDCountInMainData(idsPlusBrandSeries, mainData);

            AddNewFitmUpdateCountInMainData(brandKeyPlusSKuFitmentUpdate, mainData);

            AddTotalSkuCountInMainData(skuPlusBrandsSeries, mainData);

            AddProblemSkuCountInMainData(problematicBrandSeriesPlusSKU, mainData);
        }
        private void AddProblematicInfo(List<string> exportLink, int problematicSKUPosition)
        {
            if (exportLink[problematicSKUPosition] != "")
            {
                string dataLine = string.Empty;
                if (exportLink[problematicSKUPosition] == "3")
                {
                    dataLine = exportLink[1] + "|" + exportLink[2] + "|Live|Missing fitment information / Missing images|Request was sent / Task to designers was sent";
                }
                else if (exportLink[problematicSKUPosition] == "1")
                {
                    dataLine = exportLink[1] + "|" + exportLink[2] + "|Live|Missing fitment information|Request was sent";
                }
                else
                {
                    dataLine = exportLink[1] + "|" + exportLink[2] + "|Live|Missing images|Task to designers was sent";
                }
                problematicInfo.Add(dataLine);
            }
        }
        private void GenerateProblematicSKuList()
        {
            problematicSKU = new string[problematicInfo.Count + 1, 5];

            problematicSKU[0, 0] = "Brand";
            problematicSKU[0, 1] = "SKU";
            problematicSKU[0, 2] = "Status(Dead/Live)";
            problematicSKU[0, 3] = "Reason";
            problematicSKU[0, 4] = "What was done?";

            int position = 1;
            foreach (string s in problematicInfo)
            {
                string[] dataList = s.Split('|');

                for (int i = 0; i < dataList.Length; i++)
                {
                    problematicSKU[position, i] = dataList[i];
                }
                position++;
            }
        }
        private void AddFitmentDataInfo(List<string> exportLink, int fitmUpdatePosition)
        {
            if (exportLink[fitmUpdatePosition] == "fitment update")
            {
                string dataList = exportLink[1] + "|" + exportLink[2] + "|" + exportLink[12] + "|" + exportLink[3];
                fitmentUpdateInfo.Add(dataList);
            }
        }
        private void GenerateFitmentUpdateList()
        {
            fitmentUpdateList = new string[fitmentUpdateInfo.Count + 1, 4];

            fitmentUpdateList[0, 0] = "Brand";
            fitmentUpdateList[0, 1] = "SKU";
            fitmentUpdateList[0, 2] = "linkwww";
            fitmentUpdateList[0, 3] = "Product Name";

            int position = 1;
            foreach (string s in fitmentUpdateInfo)
            {
                bool check = true;
                string[] dataList = s.Split('|');

                for (int i = 1; i < position; i++)
                {
                    if (dataList[0] == fitmentUpdateList[i, 0] && dataList[1] == fitmentUpdateList[i, 1])
                    {
                        check = false;
                        break;
                    }
                }

                if (check == false)
                {
                    continue;
                }

                for (int i = 0; i < dataList.Length; i++)
                {
                    fitmentUpdateList[position, i] = dataList[i];
                }
                position++;
            }

        }
        private void AddNewProductInfo(List<string> exportLink, int brandKeyPosition, HashSet<string> newSKU)
        {
            if (newSKU.Contains(exportLink[brandKeyPosition]))
            {
                List<string> dataLine = new List<string>();
                dataLine.Add(exportLink[1]);
                dataLine.Add("No");
                dataLine.Add(exportLink[2]);
                dataLine.Add(exportLink[0]);
                dataLine.Add(exportLink[12]);
                dataLine.Add(exportLink[7]);
                dataLine.Add(exportLink[9]);
                dataLine.Add(exportLink[11]);
                dataLine.Add(exportLink[3]);
                dataLine.Add(exportLink[4]);

                newProductsInfo.Add(dataLine);
            }
        }
        private void AddMfrPlusSKU(List<string> exportLink, HashSet<string> newSKU, int brandKeyPosition)
        {
            mfrPlusSKU.Add("Manufacturer ID|SKU");
            if (newSKU.Contains(exportLink[brandKeyPosition]))
            {
                mfrPlusSKU.Add(exportLink[8] + "|" + exportLink[2]);
            }
        }
        private void GenerateNewSKuList()
        {
            newSKUList = new string[newProductsInfo.Count + 1, 10];

            newSKUList[0, 0] = "Brand";
            newSKUList[0, 1] = "New Series";
            newSKUList[0, 2] = "SKU";
            newSKUList[0, 3] = "Product ID";
            newSKUList[0, 4] = "linkwww";
            newSKUList[0, 5] = "Make";
            newSKUList[0, 6] = "Model";
            newSKUList[0, 7] = "Years";
            newSKUList[0, 8] = "Product Name";
            newSKUList[0, 9] = "Child Title";

            for (int i = 1; i < newSKUList.GetLength(0); i++)
            {
                for (int x = 0; x < newSKUList.GetLength(1); x++)
                {
                    newSKUList[i, x] = newProductsInfo[i - 1][x];
                }
            }

        }
        private void AddProblematicSKU(List<string> exportLink, int brandSeriesPos, int problematicPos, HashSet<string> problematicBrandSeriesPlusSKU)
        {
            if (exportLink[6].Contains("{\"not_our_mmy\":{\"unknown\":{\"unknown\":{\"0\":{") && exportLink[5].Contains("images/no-image.jpg"))
            {
                exportLink[problematicPos] = "3";
                problematicBrandSeriesPlusSKU.Add(exportLink[brandSeriesPos] + "|" + exportLink[2] + "|" + exportLink[problematicPos]);
            }
            else if (exportLink[6].Contains("{\"not_our_mmy\":{\"unknown\":{\"unknown\":{\"0\":{"))
            {
                exportLink[problematicPos] = "1";
                problematicBrandSeriesPlusSKU.Add(exportLink[brandSeriesPos] + "|" + exportLink[2] + "|" + exportLink[problematicPos]);
            }
            else if (exportLink[5].Contains("images/no-image.jpg"))
            {
                exportLink[problematicPos] = "2";
                problematicBrandSeriesPlusSKU.Add(exportLink[brandSeriesPos] + "|" + exportLink[2] + "|" + exportLink[problematicPos]);
            }
        }
        private void AddTotalWebsiteSKU(List<string> exportLink, int brandSeriesPos, HashSet<string> skuPlusBrandsSeries)
        {
            skuPlusBrandsSeries.Add(exportLink[brandSeriesPos] + "|" + exportLink[2]);
        }
        private void GetNewAndFitmentUpdateSKU(HashSet<string> newSKU, HashSet<string> fitmentUpdateSKU, int pDataBrandKeyPosition)
        {
            for (int i = 1; i < fileReader.PData.PDData1.Count; i++)
            {
                if (fileReader.PData.PDData1[i][2].ToLower() == "new")
                {
                    newSKU.Add(fileReader.PData.PDData1[i][pDataBrandKeyPosition]);
                }

                if (fileReader.PData.PDData1[i][2].ToLower() == "fitment update")
                {
                    fitmentUpdateSKU.Add(fileReader.PData.PDData1[i][pDataBrandKeyPosition]);
                }
            }
        }
        private void AddFitmentUpdate(List<string> exportLink, int brandsSeriesPos, int brandKayPos, int newIDPos, int fitmentPos, HashSet<string> brandKeyPlusSKuFitmentUpdate, HashSet<string> newSKU, HashSet<string> fitmentUpdateSKU)
        {
            if (fitmentUpdateCheck)
            {
                if (fitmentUpdateSKU.Contains(exportLink[brandKayPos]))
                {
                    exportLink[fitmentPos] = "fitment update";
                    brandKeyPlusSKuFitmentUpdate.Add(exportLink[brandsSeriesPos] + "|" + exportLink[2]);
                }
            }
            else
            {
                if (fitmentUpdateSKU.Contains(exportLink[brandKayPos]))
                {
                    exportLink[fitmentPos] = "fitment update";
                    brandKeyPlusSKuFitmentUpdate.Add(exportLink[brandsSeriesPos] + "|" + exportLink[2]);
                    return;
                }

                if (exportLink[newIDPos] == "new")
                {
                    if (!newSKU.Contains(exportLink[brandKayPos]) && !exportLink[6].Contains("{\"not_our_mmy\":{\"unknown\":{\"unknown\":{\"0\":{"))
                    {
                        exportLink[fitmentPos] = "fitment update";
                        brandKeyPlusSKuFitmentUpdate.Add(exportLink[brandsSeriesPos] + "|" + exportLink[2]);
                    }
                }
            }
            
            //else
            //{
            //    if (fileReader.ID.ProdIDMMY != null)
            //    {
            //        string idMMYKey = exportLink[0] + '|' + exportLink[7] + '|' + exportLink[9] + '|' + exportLink[11];
            //        if (!fileReader.ID.ProdIDMMY.Contains(idMMYKey))
            //        {
            //            exportLink[fitmentPos] = "fitment update";
            //            brandKeyPlusSKuFitmentUpdate.Add(exportLink[brandsSeriesPos] + "|" + exportLink[2]);
            //        }
            //    }
            //}
        }
        private void AddNewIDsCount(List<string> exportLink, int brandSeriesPos, int newIDCheckPosition, HashSet<string> idsPlusBrandSeries)
        {
            if (!fileReader.ID.ProdID.Contains(exportLink[0]))
            {
                exportLink[newIDCheckPosition] = "new";
                idsPlusBrandSeries.Add(exportLink[0] + "|" + exportLink[brandSeriesPos]);
            }
        }
        private void AddNewSKUCount(List<string> exportLink, int brandKayPos, HashSet<string> allSKU)
        {
            allSKU.Add(exportLink[brandKayPos]);
        }
        
        private void AddBrands(List<string> exportLink, int brandSeriesPos, SortedSet<string> brandSeriesKey, SortedSet<string> brands)
        {
            exportLink[brandSeriesPos] = exportLink[1].Remove(exportLink[1].Length - 1);
            brandSeriesKey.Add(exportLink[brandSeriesPos]);
            brands.Add(exportLink[1]);
        }
        private void AddCheckColumnsToExpLinksHeader(List<string> exportLink, int problematicPos, int fitmentPos, int newIDPos, int brandSeriesPos)
        {
            exportLink.Insert(problematicPos, "Problematic Check");
            exportLink.Insert(fitmentPos, "Fitment Update Check");
            exportLink.Insert(newIDPos, "New ID Check");
            exportLink.Insert(brandSeriesPos, "Brand+Series");
        }
        private void AddCheckColumnsToExpLinksEmpty(List<string> exportLink, int problematicPos, int fitmentPos, int newIDPos, int brandSeriesPos)
        {
            exportLink.Insert(problematicPos, string.Empty);
            exportLink.Insert(fitmentPos, string.Empty);
            exportLink.Insert(newIDPos, string.Empty);
            exportLink.Insert(brandSeriesPos, string.Empty);
        }
        private void GenerateHeader()
        {
            message.MessageTriger("Генерация шапки...");

            header = new string[7, 11];
            DateTime dateTime = DateTime.UtcNow.Date;
            string currentMonth = dateTime.ToString("MM");
            string currentYear = dateTime.ToString("yyyy");

            for (int i = 0; i < 6; i++)
            {
                for (int y = 0; y < 10; y++)
                {
                    header[i, y] = string.Empty;
                }
            }

            header[0, 0] = "General Info";
            header[1, 0] = "Category:";
            header[2, 0] = "Importer:";
            header[3, 0] = "Start Data:";
            header[4, 0] = "Finish Data:";
            header[5, 0] = "Elapsed Time:";
            header[6, 0] = "Scheduled Time:";
            header[1, 2] = CategoryValue[0];
            header[2, 2] = userName;
            header[3, 2] = "'" + currentMonth + ".01." + currentYear;
            header[4, 2] = "'" + dateTime.ToString("MM.dd.yyyy");
            header[5, 2] = "2 days";
            header[6, 2] = "2 days";
            header[1, 7] = "PRODUCT CATEGORY:";
            header[1, 10] = CategoryValue[1];
        }
        private void GenerateJobberApp(SortedSet<string> brands)
        {
            message.MessageTriger("Генерация джобер/апп...");

            jobberApp = new string[brands.Count + 1, 19];

            jobberApp[0, 0] = "Brand / Input Data";
            jobberApp[0, 2] = "Jobber(data)";
            jobberApp[0, 8] = "Application(data)";
            jobberApp[0, 13] = "Other used data";

            int position = 1;
            foreach (string s in brands)
            {
                jobberApp[position, 0] = s.Remove(s.Length - 1);
                for (int i = 1; i < fileReader.PData.PDData2.Count; i++)
                {
                    if (s.ToLower() == fileReader.PData.PDData2[i][0].ToLower())
                    {
                        jobberApp[position, 2] = fileReader.PData.PDData2[i][2];
                        jobberApp[position, 8] = fileReader.PData.PDData2[i][3];
                    }
                }
                position++;
            }

        }

        private void AddImgPDfVideo(int pDataBrandSeriesPosition)
        {
            for (int x = 2; x < mainData.GetLength(0); x++)
            {
                int newImageCount = 0;
                int newPDfCount = 0;
                int newVideoCount = 0;
                int number;

                for (int i = 1; i < fileReader.PData.PDData1.Count; i++)
                {
                    if (mainData[x, 1].ToLower() == fileReader.PData.PDData1[i][pDataBrandSeriesPosition].ToLower())
                    {
                        if (fileReader.PData.PDData1[i][7] != "")
                        {
                            if (Int32.TryParse(fileReader.PData.PDData1[i][7], out number))
                            {
                                newImageCount += number;
                            }
                            else
                            {
                                newImageCount++;
                            }
                        }

                        if (fileReader.PData.PDData1[i][8] != "")
                        {
                            if (Int32.TryParse(fileReader.PData.PDData1[i][8], out number))
                            {
                                newPDfCount += number;
                            }
                            else
                            {
                                newPDfCount++;
                            }
                        }

                        if (fileReader.PData.PDData1[i][9] != "")
                        {
                            if (Int32.TryParse(fileReader.PData.PDData1[i][9], out number))
                            {
                                newVideoCount += number;
                            }
                            else
                            {
                                newVideoCount++;
                            }
                        }
                    }
                }
                mainData[x, 6] = newImageCount.ToString();
                mainData[x, 7] = newPDfCount.ToString();
                mainData[x, 8] = newVideoCount.ToString();
            }
        }
        private void AddRanksAndRAPs(int pData2BrandKeyPosition)
        {
            for (int x = 2; x < mainData.GetLength(0); x++)
            {
                for (int y = 1; y < fileReader.PData.PDData2.Count; y++)
                {
                    if (mainData[x, 1].ToLower() == fileReader.PData.PDData2[y][pData2BrandKeyPosition].ToLower())
                    {
                        if (fileReader.PData.PDData2[y][4] != "")
                        {
                            mainData[x, 9] = fileReader.PData.PDData2[y][4];
                            mainData[x, 0] = fileReader.PData.PDData2[y][1];
                            break;
                        }
                        else
                        {
                            mainData[x, 9] = "0";
                            mainData[x, 0] = fileReader.PData.PDData2[y][1];
                            break;
                        }
                    }
                }
            }
        }
        private void AddImprovedStructure(int pDataBrandSeriesPosition)
        {
            for (int x = 2; x < mainData.GetLength(0); x++)
            {
                List<string> subtypesList = new List<string>();

                for (int y = 1; y < fileReader.PData.PDData1.Count; y++)
                {
                    if (fileReader.PData.PDData1[y][pDataBrandSeriesPosition].ToLower() == mainData[x, 1].ToLower())
                    {
                        subtypesList.Add(fileReader.PData.PDData1[y][6]);
                    }
                }

                string[] uniqueSubtypes = subtypesList.Distinct().ToArray();
                string resultSubtypes = String.Join("; ", uniqueSubtypes);

                for (int y = 3; y < 10; y++)
                {
                    int number;
                    if (Int32.TryParse(mainData[x, y], out number))
                    {
                        if (number > 0)
                        {
                            mainData[x, 10] = resultSubtypes;
                            break;
                        }
                    }
                }
            }
        }
        private void AddUpdateType(int pDataBrandKeyPosition)
        {
            for (int i = 2; i < mainData.GetLength(0); i++)
            {
                mainData[i, 2] = "No Changes";
                bool updateCheck = true;
                for (int y = 1; y < fileReader.PData.PDData1.Count; y++)
                {
                    if (updateCheck == false)
                    {
                        if (fileReader.PData.PDData1[y][pDataBrandKeyPosition].ToLower() == mainData[i, 1].ToLower() && fileReader.PData.PDData1[y][2] != "" && fileReader.PData.PDData1[y][2].ToLower() != "old")
                        {
                            mainData[i, 2] = "Updated part of the brand which was in PT";
                            break;
                        }
                        else
                        {
                            mainData[i, 2] = "No Changes";
                        }
                    }
                    else
                    {
                        if (fileReader.PData.PDData1[y][pDataBrandKeyPosition].ToLower() == mainData[i, 1].ToLower())
                        {
                            if (fileReader.PData.PDData1[y][2] == "")
                            {
                                updateCheck = false;
                                continue;
                            }
                            else if (fileReader.PData.PDData1[y][2] == "new")
                            {
                                mainData[i, 2] = "Uploaded part of the brand to PT";
                            }
                            else if (fileReader.PData.PDData1[y][2].ToLower() != "old")
                            {
                                mainData[i, 2] = "Updated part of the brand which was in PT";
                                break;
                            }
                        }
                    }
                }

                for (int y = 3; y < 10; y++)
                {
                    int number;
                    if (Int32.TryParse(mainData[i, y], out number))
                    {
                        if (number > 0)
                        {
                            if (mainData[i, 2] == "No Changes")
                            {
                                mainData[i, 2] = "Updated part of the brand which was in PT";
                                break;
                            }
                        }
                    }
                }
            }
        }
        private void AddHeaderToMainData()
        {
            mainData[0, 0] = "Result";
            mainData[1, 0] = "RANK";
            mainData[1, 1] = "Brand";
            mainData[1, 2] = "Update Type";
            mainData[1, 3] = "Added SKUs";
            mainData[1, 4] = "Added Product IDs";
            mainData[1, 5] = "Updated fitment data";
            mainData[1, 6] = "Added images";
            mainData[1, 7] = "Added PDF";
            mainData[1, 8] = "Added Video";
            mainData[1, 9] = "Closed Tickets/Reports";
            mainData[1, 10] = "Improved structure of parents/children products according to recommendations from PM dept";
            mainData[1, 11] = "Total website SKUs";
            mainData[1, 12] = "Problematic SKU's";
            mainData[1, 13] = "New Series";
            mainData[1, 14] = "Unsolved Problems";
            mainData[1, 15] = "Difficulties with project";
            mainData[1, 16] = "Notes";
        }
        private void GenerateChildTitleDupList()
        {
            if (fileReader.ChtTitleDuplicates == null)
                return;

            List<List<string>> chtData = new List<List<string>>();
            List<string> statBrands = new List<string>();
            List<string> statData = new List<string>();
            foreach (var row in fileReader.ChtTitleDuplicates)
            {
                if (row[1]!="" && brandSeriesKey.Contains(row[1].Remove(row[1].Length - 1)))
                {
                    List<string> list = new List<string>();
                    for (int i = 0; i < row.Count-3; i++)
                        list.Add(row[i]);

                    chtData.Add(list);
                }
                if (row[7]!="" && brandSeriesKey.Contains(row[7].Remove(row[7].Length - 1)))
                {
                    statBrands.Add(row[7].Remove(row[7].Length - 1));
                    statData.Add(row[8]);
                }
            }

            foreach (var brand in brandSeriesKey)
            {
                if (!statBrands.Contains(brand))
                {
                    //string totalSkuCount = "0";
                    //for (int i = 2; i < mainData.GetLength(0); i++)
                    //{
                    //    if (mainData[i, 1].ToLower() == brand.ToLower())
                    //    {
                    //        totalSkuCount = mainData[i, 11];
                    //        break;
                    //    }
                    //}
                    
                    string brandName = brand;
                    statBrands.Add(brandName);
                    statData.Add("0 / 0, 0.00%");
                }
            }
            if (chtData.Count < statBrands.Count)
            {
                chtTitleDuplicates = new string[statBrands.Count + 1, 9];
            } else
            {
                chtTitleDuplicates = new string[chtData.Count + 1, 9];
            }

            chtTitleDuplicates[0, 0] = "Product Name";
            chtTitleDuplicates[0, 1] = "Brand";
            chtTitleDuplicates[0, 2] = "Make";
            chtTitleDuplicates[0, 3] = "Model";
            chtTitleDuplicates[0, 4] = "Years";
            chtTitleDuplicates[0, 5] = "Parent IDs with Child title Duplication";
            chtTitleDuplicates[0, 6] = "";
            chtTitleDuplicates[0, 7] = "Brands";
            chtTitleDuplicates[0, 8] = "Statistic (all Parents per Brand / parents with child title duplication, %)";

            if (chtData.Count > 0)
            {
                for (int i = 1; i < chtData.Count; i++)
                {
                    for (int x = 0; x < chtData[0].Count; x++)
                        chtTitleDuplicates[i, x] = chtData[i - 1][x];
                }
            }

            for (int x = 1; x < statBrands.Count + 1; x++)
            {
                if (statBrands[x - 1] == "Torxe")
                {
                    chtTitleDuplicates[x, 7] = statBrands[x - 1] + "™";
                } else
                {
                    chtTitleDuplicates[x, 7] = statBrands[x - 1] + "®";
                }
                chtTitleDuplicates[x, 8] = statData[x-1];
            }
        }
        private string[] GetCategoryValue()
        {
            string[] subtypes = new string[fileReader.PData.PDData1.Count - 1];
            string[] categoryNames = new string[fileReader.PData.PDData1.Count - 1];

            for (int i = 0; i < subtypes.Length; i++)
            {
                subtypes[i] = fileReader.PData.PDData1[i + 1][6];
                categoryNames[i] = fileReader.PData.PDData1[i + 1][5];
            }

            string[] categoryValue = new string[3];
            string[] uniqueCategoryNames = categoryNames.Distinct().ToArray();
            string[] uniqueSubtypes = subtypes.Distinct().ToArray();

            categoryValue[0] = string.Join("; ", uniqueCategoryNames);
            categoryValue[1] = string.Join("; ", uniqueSubtypes);
            categoryValue[2] = fileReader.PData.PDData1[2][5];

            return categoryValue;
        }


        private string[] categoryValue;
        private string[,] mainData;
        private string[,] header;
        private string[,] jobberApp;
        private string[,] newSKUList;
        private string[,] fitmentUpdateList;
        private string[,] problematicSKU;
        private string[,] chtTitleDuplicates;
        private HashSet<string> mfrPlusSKU = new HashSet<string>();
        private List<List<string>> newProductsInfo = new List<List<string>>();
        private HashSet<string> fitmentUpdateInfo = new HashSet<string>();
        private HashSet<string> problematicInfo = new HashSet<string>();

        public string[] CategoryValue
        {
            get
            {
                if (categoryValue == null)
                {
                    categoryValue = GetCategoryValue();
                }
                return categoryValue;
            }
        }
        public string[,] MainData
        {
            get
            {
                if (mainData == null)
                {
                    GenerateMainData();
                }
                return mainData;
            }
        }
        public string[,] Header
        {
            get
            {
                if (header == null)
                {
                    GenerateHeader();
                }
                return header;
            }
        }
        public string[,] JobberApp
        {
            get
            {
                if (jobberApp == null)
                {
                    GenerateMainData();
                }
                return jobberApp;
            }
        }
        public string[,] NewSKUList
        {
            get
            {
                if (newSKUList == null)
                {
                    GenerateNewSKuList();
                }
                return newSKUList;
            }
        }
        public string[,] FitmentUpdateList
        {
            get
            {
                if (fitmentUpdateList == null)
                {
                    GenerateFitmentUpdateList();
                }
                return fitmentUpdateList;
            }
        }
        public string[,] ProblematicSKU
        {
            get
            {
                if (problematicSKU == null)
                {
                    GenerateProblematicSKuList();
                }
                return problematicSKU;
            }
        }
        public string[,] ChildTitleDuplicates
        {
            get
            {
                if (chtTitleDuplicates == null)
                {
                    GenerateChildTitleDupList();
                }
                return chtTitleDuplicates;
            }
        }
        public HashSet<string> MfrPlusSKU
        {
            get
            {
                if (mfrPlusSKU == null)
                {
                    GenerateNewSKuList();
                }
                return mfrPlusSKU;
            }
        }
    }
}
