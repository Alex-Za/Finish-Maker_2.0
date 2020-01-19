using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Threading;
using System.Collections.ObjectModel;
using System.Windows.Media;
using System.Windows.Threading;
using System.Diagnostics;
using Microsoft.Win32;

namespace Finish_Maker_Demo
{
    class FinishMakerViewModel : INotifyPropertyChanged
    {

        FinishMakerModel finishMaker = new FinishMakerModel();
        public ObservableCollection<ExportLinks> ExpLinksList { get; set; }
        public ObservableCollection<PD> PDList { get; set; }
        public ObservableCollection<ID> IDList { get; set; }
        public ObservableCollection<ChildTitleDuplicates> ChtDuplicatesList { get; set; }

        string[] ghbrjk = { "Sidorchenko", "Sidorchuk", "Sidorkidze", "Sidorkisyan", "Van Der Sidorkin", "Sidorkopulos", "Sidorenko", "Sidorkov", "Sidorski", "Sidormen", "Sidorkinyo", "Sidorishkin", "Sidorkyauskas" };

        public FinishMakerViewModel()
        {
            ExpLinksList = new ObservableCollection<ExportLinks>();
            PDList = new ObservableCollection<PD>();
            IDList = new ObservableCollection<ID>();
            ChtDuplicatesList = new ObservableCollection<ChildTitleDuplicates>();
            finishMaker.ProductDataCheck = true;
            finishMaker.ValidateFiles = true;
            UserName = "Evgeniy " + ghbrjk[new Random().Next(0, ghbrjk.Length)];

            worker = new BackgroundWorker();
            worker.WorkerReportsProgress = true;
            worker.DoWork += worker_DoWork;
            worker.ProgressChanged += worker_ProgressChanged;
        }

        private RelayCommand addExpLinksCommand;
        private RelayCommand addPDCommand;
        private RelayCommand addIDCommand;
        private RelayCommand addChtDuplicatesCommand;
        private RelayCommand start;
        private RelayCommand deleteCommand;
        private ConsoleText consoleTextProperty;
        public RelayCommand AddExpLinksCommand
        {
            get
            {
                return addExpLinksCommand ??
                  (addExpLinksCommand = new RelayCommand(obj =>
                  {
                      OpenFileDialog openFileDialog = new OpenFileDialog();
                      if (openFileDialog.ShowDialog() == true)
                      {
                          ExportLinks exportLinks = new ExportLinks();
                          exportLinks.Path = openFileDialog.FileName;
                          exportLinks.ViewPath = openFileDialog.FileName.Substring(openFileDialog.FileName.LastIndexOf("\\") + 1);
                          exportLinks.ID = ExpLinksList.Count + 1;
                          ExpLinksList.Add(exportLinks);
                      }
                  }));
            }
        }
        public RelayCommand AddPDCommand
        {
            get
            {
                return addPDCommand ??
                  (addPDCommand = new RelayCommand(obj =>
                  {
                      OpenFileDialog openFileDialog = new OpenFileDialog();
                      if (openFileDialog.ShowDialog() == true)
                      {
                          PD pdLinks = new PD();
                          pdLinks.Path = openFileDialog.FileName;
                          pdLinks.ViewPath = openFileDialog.FileName.Substring(openFileDialog.FileName.LastIndexOf("\\") + 1);
                          pdLinks.ID = PDList.Count + 1;
                          PDList.Add(pdLinks);
                      }
                  }));
            }
        }
        public RelayCommand AddIDCommand
        {
            get
            {
                return addIDCommand ??
                  (addIDCommand = new RelayCommand(obj =>
                  {
                      OpenFileDialog openFileDialog = new OpenFileDialog();
                      if (openFileDialog.ShowDialog() == true)
                      {
                          ID idList = new ID();
                          idList.Path = openFileDialog.FileName;
                          idList.ViewPath = openFileDialog.FileName.Substring(openFileDialog.FileName.LastIndexOf("\\") + 1);
                          idList.ID = IDList.Count + 1;
                          IDList.Add(idList);
                      }
                  }));
            }
        }
        public RelayCommand AddChtDuplicatesCommand
        {
            get
            {
                return addChtDuplicatesCommand ??
                  (addChtDuplicatesCommand = new RelayCommand(obj =>
                  {
                      OpenFileDialog openFileDialog = new OpenFileDialog();
                      if (openFileDialog.ShowDialog() == true)
                      {
                          ChildTitleDuplicates chtList = new ChildTitleDuplicates();
                          chtList.Path = openFileDialog.FileName;
                          chtList.ViewPath = openFileDialog.FileName.Substring(openFileDialog.FileName.LastIndexOf("\\") + 1);
                          chtList.ID = ChtDuplicatesList.Count + 1;
                          ChtDuplicatesList.Add(chtList);
                      }
                  }));
            }
        }
        public bool IsSelectedExpLinkCheck
        {
            get { return finishMaker.ExportLinkCheck; }
            set
            {
                finishMaker.ExportLinkCheck = value;

                if (finishMaker.ProductDataCheck == value)
                {
                    IsSelectedPDCheck = !value;
                }

                OnPropertyChanged("IsSelectedExpLinkCheck");
            }
        }
        public bool IsSelectedPDCheck
        {
            get { return finishMaker.ProductDataCheck; }
            set
            {
                finishMaker.ProductDataCheck = value;

                if (finishMaker.ExportLinkCheck == value)
                {
                    IsSelectedExpLinkCheck = !value;
                }
                
                OnPropertyChanged("IsSelectedPDCheck");
            }
        }
        public bool ValidateFiles
        {
            get { return finishMaker.ValidateFiles; }
            set
            {
                finishMaker.ValidateFiles = value;
                OnPropertyChanged("ValidateFiles");
            }
        }
        public RelayCommand Start
        {
            get
            {
                return start ??
                  (start = new RelayCommand(obj =>
                  {
                      worker.RunWorkerAsync();
                  }));
            }
        }

        public RelayCommand DeleteCommand
        {
            get
            {
                return deleteCommand ?? (deleteCommand = new RelayCommand(obj =>
                {
                    Files currentFile = obj as Files;
                    if (currentFile != null)
                    {

                        if (currentFile is ExportLinks)
                        {
                            ExpLinksList.Remove(currentFile as ExportLinks);
                        }
                        else if (currentFile is PD)
                        {
                            PDList.Remove(currentFile as PD);
                        }
                        else if (currentFile is ID)
                        {
                            IDList.Remove(currentFile as ID);
                        }
                        else if (currentFile is ChildTitleDuplicates)
                        {
                            ChtDuplicatesList.Remove(currentFile as ChildTitleDuplicates);
                        }

                    }
                }));
            }
        }
        public string UserName
        {
            get { return finishMaker.UserName; }
            set
            {
                finishMaker.UserName = value;
                OnPropertyChanged("UserName");
            }
        }

        public ConsoleText ConsoleTextProperty
        {
            get { return consoleTextProperty; }
            set
            {
                consoleTextProperty = value;
                OnPropertyChanged("ConsoleTextProperty");
            }
        }
        public int Progress
        {
            get
            {
                return finishMaker.Progress;
            }
            set
            {
                finishMaker.Progress = value;
                OnPropertyChanged("Progress");
            }
        }

        private BackgroundWorker worker;

        private void changeProgress(int count)
        {
            this.worker.ReportProgress(count);
        }

        private void worker_DoWork(object sender, DoWorkEventArgs e)
        {
            OpenFileDialog fd = new OpenFileDialog();
            fd.ValidateNames = false;
            fd.CheckFileExists = false;
            fd.CheckPathExists = true;
            fd.FileName = "Folder Selection.";

            if (fd.ShowDialog() == true)
            {
                string saveFilePath = Path.GetDirectoryName(fd.FileName);

                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();

                ConsoleTextProperty = new ConsoleText { TheText = "In progress..." + Environment.NewLine, TheColor = Brushes.Black };

                List<List<string>> filePath = GetAllPathes(ExpLinksList, PDList, IDList, ChtDuplicatesList);

                FileReader fileReader = new FileReader(filePath, IsSelectedPDCheck);

                changeProgress(10);
                Mistakes mistakesCheck = new Mistakes(fileReader, filePath, IsSelectedPDCheck);

                if (ValidateFiles == true)
                {
                    if (mistakesCheck.CriticalErrors != null)
                    {
                        ConsoleText someConsoleText = new ConsoleText { TheColor = Brushes.Red, TheText = mistakesCheck.CriticalErrors };
                        ConsoleTextProperty = someConsoleText;
                        changeProgress(0);
                        return;
                    }

                    if (mistakesCheck.OtherErrors != null)
                    {
                        ConsoleText someConsoleText = new ConsoleText { TheColor = Brushes.Red, TheText = mistakesCheck.OtherErrors };
                        ConsoleTextProperty = someConsoleText;
                    }
                }

                Processing processing = new Processing(fileReader, UserName);
                Writer writer = new Writer(processing, changeProgress, saveFilePath);

                writer.Write();

                //writer.WriteExcelFile("C:\\Programming\\C#\\Test.xlsx", fileReader.ID.ProdIDMMY);

                if (ValidateFiles == true)
                {
                    if (mistakesCheck.OtherErrors != null)
                    {
                        ConsoleText someConsoleText = ConsoleTextProperty;
                        someConsoleText.TheText += "Done";
                        ConsoleTextProperty = someConsoleText;
                    }
                    else
                    {
                        ConsoleText someConsoleText = new ConsoleText { TheColor = Brushes.Black, TheText = "Done" };
                        ConsoleTextProperty = someConsoleText;
                    }
                }
                else
                {
                    ConsoleText someConsoleText = new ConsoleText { TheColor = Brushes.Black, TheText = "Done" };
                    ConsoleTextProperty = someConsoleText;
                }

                stopWatch.Stop();
                TimeSpan ts = stopWatch.Elapsed;

                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}", ts.Hours, ts.Minutes, ts.Seconds);

                ConsoleText workingTime = ConsoleTextProperty;
                workingTime.TheText += Environment.NewLine + "Время работы программы: " + elapsedTime;
                ConsoleTextProperty = workingTime;
                }
            
        }

        void worker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Progress = e.ProgressPercentage;
        }

        private List<List<string>> GetAllPathes(ObservableCollection<ExportLinks> expLinkPath, ObservableCollection<PD> pdPath, ObservableCollection<ID> idPath, ObservableCollection<ChildTitleDuplicates> chtPath)
        {
            var expLinkNames = from exp in expLinkPath select exp.Path;
            var pdNames = from pd in pdPath select pd.Path;
            var idNames = from id in idPath select id.Path;
            var chtNames = from cht in chtPath select cht.Path;
            List<List<string>> result = new List<List<string>>();
            result.Add(expLinkNames.ToList());
            result.Add(pdNames.ToList());
            result.Add(idNames.ToList());
            result.Add(chtNames.ToList());
            return result;
        }


        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName]string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
    }
}
