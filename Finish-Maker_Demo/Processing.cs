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
        private List<List<string>> exportLinks;
        public Processing(FileReader reader, string userName)
        {
            fileReader = reader;
            this.userName = userName;
        }

        private void GenerateMainData()
        {
            exportLinks = fileReader.ExportLinks.ToList();

            int problematicSKUPosition = exportLinks[0].Count - 1;
            int fitmUpdatePosition = exportLinks[0].Count;
            int newIDCheckPosition = exportLinks[0].Count + 1;
            int brandSeriesPosition = exportLinks[0].Count + 2;
            int brandKayPosition = exportLinks[0].Count + 3;

            AddCheckColumnsToExpLinks(exportLinks, problematicSKUPosition, fitmUpdatePosition, newIDCheckPosition, brandSeriesPosition);
            AddBrands(exportLinks, brandSeriesPosition);
            AddNewSKUCount(exportLinks, brandKayPosition);
            AddNewIDsCount(exportLinks, brandSeriesPosition, newIDCheckPosition);
            AddFitmentUpdate(exportLinks, brandSeriesPosition, brandKayPosition, newIDCheckPosition, fitmUpdatePosition);
            AddImgPDfVideo();
            AddRanksAndRAPs();
            AddImprovedStructure();
            AddUpdateType();
            AddTotalWebsiteSKU(exportLinks, brandSeriesPosition);
            AddProblematicSKU(exportLinks, brandSeriesPosition, problematicSKUPosition);
            AddHeaderToMainData();
        }
        private void GenerateHeader()
        {
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
            header[3, 2] = currentMonth + ".01." + currentYear;
            header[4, 2] = dateTime.ToString("MM/dd/yyyy");
            header[5, 2] = "2 days";
            header[6, 2] = "2 days";
            header[1, 7] = "PRODUCT CATEGORY:";
            header[1, 10] = CategoryValue[1];
        }
        private void GenerateJobberApp()
        {
            SortedSet<string> brands = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 1; i < fileReader.PData.PDData1.Count; i++)
            {
                brands.Add(fileReader.PData.PDData1[i][0]);
            }

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

        private void AddCheckColumnsToExpLinks(List<List<string>> exportLinks, int problematicPos, int fitmentPos, int newIDPos, int brandSeriesPos)
        {
            exportLinks[0].Insert(problematicPos, "Problematic Check");
            exportLinks[0].Insert(fitmentPos, "Fitment Update Check");
            exportLinks[0].Insert(newIDPos, "New ID Check");
            exportLinks[0].Insert(brandSeriesPos, "Brand+Series");

            for (int i = 1; i < exportLinks.Count; i++)
            {
                exportLinks[i].Insert(problematicPos, string.Empty);
                exportLinks[i].Insert(fitmentPos, string.Empty);
                exportLinks[i].Insert(newIDPos, string.Empty);
                exportLinks[i].Insert(brandSeriesPos, string.Empty);
            }

            this.exportLinks = exportLinks;
        }
        private void AddBrands(List<List<string>> exportLinks, int brandSeriesPos)
        {
            //int pDataBrandSeriesPosition = fileReader.PData.PDData1[0].Count - 2;
            //int pDataBrandKayPosition = fileReader.PData.PDData1[0].Count - 1;

            SortedSet<string> brandSeriesKey = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);

            //добавление колонки бренд + серия в експорт линки
            for (int i = 1; i < exportLinks.Count; i++)
            {
                //if (exportLinks[i][1] == "Standard®")
                //{
                //    for (int x = 1; x < fileReader.PData.PDData1.Count; x++)
                //    {
                //        if (fileReader.PData.PDData1[x][pDataBrandKayPosition] == exportLinks[i][brandKayPosition])
                //        {
                //            exportLinks[i][brandSeries] = fileReader.PData.PDData1[x][pDataBrandSeriesPosition];
                //            break;
                //        }
                //    }
                //}
                //else
                //{
                //    exportLinks[i][brandSeries] = exportLinks[i][1].Remove(exportLinks[i][1].Length - 1);
                //}
                exportLinks[i][brandSeriesPos] = exportLinks[i][1].Remove(exportLinks[i][1].Length - 1);
                brandSeriesKey.Add(exportLinks[i][brandSeriesPos]);
            }

            mainData = new string[brandSeriesKey.Count + 2, 19];
            int positionForMainData = 2;

            foreach (string s in brandSeriesKey)
            {
                mainData[positionForMainData, 1] = s;
                positionForMainData++;

            }

            this.exportLinks = exportLinks;
        }
        private void AddNewSKUCount(List<List<string>> exportLinks, int brandKayPos)
        {
            int pDataBrandSeriesPosition = fileReader.PData.PDData1[0].Count - 2;
            int pDataBrandKayPosition = fileReader.PData.PDData1[0].Count - 1;

            HashSet<string> allSKU = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 1; i < exportLinks.Count - 1; i++)
            {
                allSKU.Add(exportLinks[i][brandKayPos]);
            }

            for (int i = 2; i < mainData.GetLength(0); i++)
            {
                int newProductCount = 0;
                for (int y = 1; y < fileReader.PData.PDData1.Count; y++)
                {
                    if (mainData[i, 1].ToLower() == fileReader.PData.PDData1[y][pDataBrandSeriesPosition].ToLower() && fileReader.PData.PDData1[y][2].ToLower() == "new" && allSKU.Contains(fileReader.PData.PDData1[y][pDataBrandKayPosition]))
                    {
                        newProductCount++;
                    }
                }
                mainData[i, 3] = newProductCount.ToString();
            }
        }
        private void AddNewIDsCount(List<List<string>> exportLinks, int brandSeriesPos, int newIDPos)
        {
            HashSet<string> idsPlusBrandSeries = new HashSet<string>();

            for (int i = 1; i < exportLinks.Count; i++)
            {
                if (!fileReader.ID.Contains(exportLinks[i][0]))
                {
                    exportLinks[i][newIDPos] = "new";
                    idsPlusBrandSeries.Add(exportLinks[i][0] + "|" + exportLinks[i][brandSeriesPos]);
                    continue;
                }
                exportLinks[i][newIDPos] = "old";
            }

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

            this.exportLinks = exportLinks;
        }
        private void AddFitmentUpdate(List<List<string>> exportLinks, int brandsSeriesPos, int brandKayPos, int newIDPos, int fitmentPos)
        {
            int pDataBrandKayPosition = fileReader.PData.PDData1[0].Count - 1;
            HashSet<string> newSKU = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int i = 1; i < fileReader.PData.PDData1.Count; i++)
            {
                if (fileReader.PData.PDData1[i][2].ToLower() == "new")
                {
                    newSKU.Add(fileReader.PData.PDData1[i][pDataBrandKayPosition]);
                }
            }

            HashSet<string> brandKeyPlusSKuFitmentUpdate = new HashSet<string>();

            for (int i = 1; i < exportLinks.Count; i++)
            {
                if (exportLinks[i][newIDPos] == "new")
                {
                    if (!newSKU.Contains(exportLinks[i][brandKayPos]) && !exportLinks[i][6].Contains("{\"not_our_mmy\":{\"unknown\":{\"unknown\":{\"0\":{"))
                    {
                        exportLinks[i][fitmentPos] = "fitment update";
                        brandKeyPlusSKuFitmentUpdate.Add(exportLinks[i][brandsSeriesPos] + "|" + exportLinks[i][2]);
                    }
                    continue;
                }

                exportLinks[i][fitmentPos] = string.Empty;
            }

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

            this.exportLinks = exportLinks;
        }
        private void AddImgPDfVideo()
        {
            int pDataBrandSeriesPosition = fileReader.PData.PDData1[0].Count - 2;

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
                            if (Int32.TryParse(fileReader.PData.PDData1[i][7], out number))
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
                            if (Int32.TryParse(fileReader.PData.PDData1[i][7], out number))
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
        private void AddRanksAndRAPs()
        {
            int pDataKeyPosition = fileReader.PData.PDData2[0].Count - 1;

            for (int x = 2; x < mainData.GetLength(0); x++)
            {
                for (int y = 1; y < fileReader.PData.PDData2.Count; y++)
                {
                    if (mainData[x, 1].ToLower() == fileReader.PData.PDData2[y][pDataKeyPosition].ToLower())
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
        private void AddImprovedStructure()
        {
            int pDataBrandSeriesPosition = fileReader.PData.PDData1[0].Count - 2;
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

                for (int y = 3; y < 8; y++)
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
        private void AddUpdateType()
        {
            int pDataKeyPosition = fileReader.PData.PDData2[0].Count - 1;

            for (int i = 2; i < mainData.GetLength(0); i++)
            {
                mainData[i, 2] = "No Changes";
                bool updateCheck = true;
                for (int y = 1; y < fileReader.PData.PDData1.Count; y++)
                {
                    if (updateCheck == false)
                    {
                        if (fileReader.PData.PDData1[y][pDataKeyPosition].ToLower() == mainData[i, 1].ToLower() && fileReader.PData.PDData1[y][2] != "" && fileReader.PData.PDData1[y][2].ToLower() != "old")
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
                        if (fileReader.PData.PDData1[y][pDataKeyPosition].ToLower() == mainData[i, 1].ToLower())
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

                for (int y = 3; y < 8; y++)
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
        private void AddTotalWebsiteSKU(List<List<string>> exportLinks, int brandSeriesPos)
        {
            HashSet<string> skuPlusBrandsSeries = new HashSet<string>();

            for (int i = 1; i < exportLinks.Count; i++)
            {
                skuPlusBrandsSeries.Add(exportLinks[i][brandSeriesPos] + '|' + exportLinks[i][2]);
            }

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
        private void AddProblematicSKU(List<List<string>> exportLinks, int brandSeriesPos, int problematicPos)
        {
            HashSet<string> problematicBrandSeriesPlusSKU = new HashSet<string>();

            for (int i = 1; i < exportLinks.Count; i++)
            {
                if (exportLinks[i][6].Contains("{\"not_our_mmy\":{\"unknown\":{\"unknown\":{\"0\":{") && exportLinks[i][5].Contains("images/no-image.jpg"))
                {
                    exportLinks[i][problematicPos] = "3";
                    problematicBrandSeriesPlusSKU.Add(exportLinks[i][brandSeriesPos] + '|' + exportLinks[i][2] + '|' + exportLinks[i][problematicPos]);

                }
                else if (exportLinks[i][6].Contains("{\"not_our_mmy\":{\"unknown\":{\"unknown\":{\"0\":{"))
                {
                    exportLinks[i][problematicPos] = "1";
                    problematicBrandSeriesPlusSKU.Add(exportLinks[i][brandSeriesPos] + '|' + exportLinks[i][2] + '|' + exportLinks[i][problematicPos]);
                }
                else if (exportLinks[i][5].Contains("images/no-image.jpg"))
                {
                    exportLinks[i][problematicPos] = "2";
                    problematicBrandSeriesPlusSKU.Add(exportLinks[i][brandSeriesPos] + '|' + exportLinks[i][2] + '|' + exportLinks[i][problematicPos]);
                }
            }

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

            this.exportLinks = exportLinks;

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
        private void GenerateNewSKuList()
        {
            int brandKayPosition = 17;
            int pDataBranKeyPosition = fileReader.PData.PDData1[0].Count - 1;
            HashSet<string> newSKU = new HashSet<string>();

            for (int i = 1; i < fileReader.PData.PDData1.Count; i++)
            {
                if (fileReader.PData.PDData1[i][2].ToLower() == "new")
                {
                    newSKU.Add(fileReader.PData.PDData1[i][pDataBranKeyPosition]);
                }
            }

            List<List<string>> newProductsInfo = new List<List<string>>();

            mfrPlusSKU = new HashSet<string>();
            mfrPlusSKU.Add("Manufacturer ID|SKU");

            for (int i = 1; i < exportLinks.Count; i++)
            {
                if (newSKU.Contains(exportLinks[i][brandKayPosition]))
                {
                    mfrPlusSKU.Add(exportLinks[i][8] + '|' + exportLinks[i][2]);

                    List<string> dataLine = new List<string>();
                    dataLine.Add(exportLinks[i][1]);
                    dataLine.Add("No");
                    dataLine.Add(exportLinks[i][2]);
                    dataLine.Add(exportLinks[i][0]);
                    dataLine.Add(exportLinks[i][12]);
                    dataLine.Add(exportLinks[i][7]);
                    dataLine.Add(exportLinks[i][9]);
                    dataLine.Add(exportLinks[i][11]);
                    dataLine.Add(exportLinks[i][3]);
                    dataLine.Add(exportLinks[i][4]);

                    newProductsInfo.Add(dataLine);
                }
            }

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
        private void GenerateFitmentUpdateList()
        {
            int fitmentUpdatePosition = 14;

            HashSet<string> fitmentUpdateSet = new HashSet<string>();

            for (int i = 1; i < exportLinks.Count; i++)
            {
                if (exportLinks[i][fitmentUpdatePosition] == "fitment update")
                {
                    string dataList = exportLinks[i][1] + '|' + exportLinks[i][2] + '|' + exportLinks[i][12] + '|' + exportLinks[i][3];

                    fitmentUpdateSet.Add(dataList);
                }
            }

            fitmentUpdateList = new string[fitmentUpdateSet.Count + 1, 4];

            fitmentUpdateList[0, 0] = "Brand";
            fitmentUpdateList[0, 1] = "SKU";
            fitmentUpdateList[0, 2] = "linkwww";
            fitmentUpdateList[0, 3] = "Product Name";

            int position = 1;
            foreach (string s in fitmentUpdateSet)
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
        private void GenerateProblematicSKuList()
        {
            int problematicSKuPosition = 13;

            HashSet<string> problematicSKuSet = new HashSet<string>();

            for (int i = 1; i < exportLinks.Count; i++)
            {
                if (exportLinks[i][problematicSKuPosition] != "")
                {
                    string dataList = string.Empty;

                    if (exportLinks[i][problematicSKuPosition] == "3")
                    {
                        dataList = exportLinks[i][1] + '|' + exportLinks[i][2] + '|' + "Live" + '|' + "Missing fitment information / Missing images" + '|' + "Request was sent / Task to designers was sent";

                    }
                    else if (exportLinks[i][problematicSKuPosition] == "1")
                    {
                        dataList = exportLinks[i][1] + '|' + exportLinks[i][2] + '|' + "Live" + '|' + "Missing fitment information" + '|' + "Request was sent";
                    }
                    else
                    {
                        dataList = exportLinks[i][1] + '|' + exportLinks[i][2] + '|' + "Live" + '|' + "Missing images" + '|' + "Task to designers was sent";
                    }

                    problematicSKuSet.Add(dataList);

                }
            }

            problematicSKU = new string[problematicSKuSet.Count + 1, 5];

            problematicSKU[0, 0] = "Brand";
            problematicSKU[0, 1] = "SKU";
            problematicSKU[0, 2] = "Status(Dead/Live)";
            problematicSKU[0, 3] = "Reason";
            problematicSKU[0, 4] = "What was done?";

            int position = 1;
            foreach (string s in problematicSKuSet)
            {
                string[] dataList = s.Split('|');

                for (int i = 0; i < dataList.Length; i++)
                {
                    problematicSKU[position, i] = dataList[i];
                }
                position++;
            }

        }
        private void GenerateChildTitleDupList()
        {
            if (fileReader.ChtTitleDuplicates == null)
            {
                return;
            }

            chtTitleDuplicates = new string[fileReader.ChtTitleDuplicates.Count, fileReader.ChtTitleDuplicates[0].Count];

            for (int i = 0; i < chtTitleDuplicates.GetLength(0); i++)
            {
                for (int x = 0; x < chtTitleDuplicates.GetLength(1); x++)
                {
                    chtTitleDuplicates[i, x] = fileReader.ChtTitleDuplicates[i][x];
                }
            }
        }
        private string[] GetCategoryValue()
        {
            string[] subtypes = new string[fileReader.PData.PDData1.Count - 1];

            for (int i = 0; i < subtypes.Length; i++)
            {
                subtypes[i] = fileReader.PData.PDData1[i + 1][6];
            }

            string[] categoryValue = new string[2];
            string[] uniqueSubtypes = subtypes.Distinct().ToArray();

            categoryValue[0] = fileReader.PData.PDData1[1][5];
            categoryValue[1] = string.Join("; ", uniqueSubtypes);

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
        private HashSet<string> mfrPlusSKU;

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
                    GenerateJobberApp();
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
