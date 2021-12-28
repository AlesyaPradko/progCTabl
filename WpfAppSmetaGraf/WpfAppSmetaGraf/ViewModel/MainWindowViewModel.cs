using System.Collections.ObjectModel;
using WpfAppSmetaGraf.Model;
using System.Windows.Input;
using System.ComponentModel;
using System.Collections.Generic;
using WpfAppSmetaGraf.Infrastructure;
using Microsoft.Win32;
using System.Windows;
using System;
using System.Threading.Tasks;

namespace WpfAppSmetaGraf.ViewModel
{
    public class MainWindowViewModel : ViewModelBase
    {
        public ModelWork _modelWork=new ModelWork();
       
        public string Adress { get; set; }
        public int IndexFontSize { get; set; }
        public int IndexAmountWorker { get; set; }
        public int IndexAmountDays { get; set; }
        public int IndexColors { get; set; }
        public DateTime DateStart { get; set; }
        

        public string Data { get; set; }
        private bool _testE=false;
        private bool _testT=false;
        private bool _testDay = false;
        private bool _testPeople=false;
        private bool _flag=true;
        private bool _exitOperation=false;
        private bool _daysOrPeople = false;
        private int _colorGet;

        public string TextErrorSm { get { return _modelWork.TextError; } set { _modelWork.TextError = value; OnPropertyChanged("TextErrorSm"); } }
        public string TextErrorGr { get { return _modelWork.TextError; } set { _modelWork.TextError = value; OnPropertyChanged("TextErrorGr"); } }
        public List<int> AmountWorkers
        {
            get { return _modelWork.AmountWorker; }
            set
            {
                _modelWork.AmountWorker = value;               
                OnPropertyChanged("AmountWorkers");              
            }
        }

        public List<int> AmountDays
        {
            get { return _modelWork.AmountWorkDays; }
            set
            {
                _modelWork.AmountWorkDays = value;
                OnPropertyChanged("AmountDays");
            }
        }

        RelayCommand _addDressFolder;
        public ICommand AddFolder
        {
            get
            {
                if (_addDressFolder == null)
                    _addDressFolder = new RelayCommand(ExecuteAddFolderCommand, CanExecuteAddFolderCommand);
                return _addDressFolder;
            }
        }

        public void ExecuteAddFolderCommand(object parameter)
        {
            System.Windows.Forms.FolderBrowserDialog openFileDlg = new System.Windows.Forms.FolderBrowserDialog();
            var result = openFileDlg.ShowDialog();
            if (result.ToString() != string.Empty)
            {
                Adress = openFileDlg.SelectedPath;
                switch (parameter.ToString()) 
                {
                    case "ChangeFolderSmeta":
                        _modelWork.AdressSmeta = Adress;  break;
                    case "ChangeFolderKS":
                        _modelWork.AdressAktKS = Adress; break;
                    case "SaveCopySmeta":
                        _modelWork.AdressSaveSmeta = Adress;  break;
                    case "SaveGraph":
                        _modelWork.AdressSaveGraph = Adress;  break;
                }                 
            }
        }
        
        public bool CanExecuteAddFolderCommand(object parameter)
        {
            if (parameter.ToString()== "ChangeFolderSmeta"|| parameter.ToString() == "ChangeFolderKS" 
                || parameter.ToString() == "SaveCopySmeta"|| parameter.ToString() == "SaveGraph")
                return true;
           else return false;
        }

        RelayCommand _addDressSmeta;
        public ICommand AddSmeta
        {
            get
            {
                if (_addDressSmeta == null)
                    _addDressSmeta = new RelayCommand(ExecuteAddSmetaCommand, CanExecuteAddSmetaCommand);
                return _addDressSmeta;
            }
        }
        public void ExecuteAddSmetaCommand(object parameter)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Files Excels(*.xlsx;*.csv)|*.xlsx;*.csv";
            if (dlg.ShowDialog() == true)
            {
                _modelWork.AdressOneSmeta = dlg.FileName;              
            }
        }
        public bool CanExecuteAddSmetaCommand(object parameter)
        {
            if (parameter.ToString() == "ChangeSmeta")
                return true;
            else return false;
        }

        RelayCommand _addSelectMod;
        public ICommand selectMod
        {
            get
            {
                if (_addSelectMod == null)
                    _addSelectMod = new RelayCommand(ExecuteAddSelectCommand, CanExecuteAddSelectCommand);
                return _addSelectMod;
            }
        }

        public void ExecuteAddSelectCommand(object parameter)
        {
            switch(IndexFontSize)
            {
                case -1:
                    _modelWork.FrontSize = 8;break;
                case 0:
                    _modelWork.FrontSize = 8; break;
                case 1:
                    _modelWork.FrontSize = 9; break;
                case 2:
                    _modelWork.FrontSize = 10; break;
                case 3:
                    _modelWork.FrontSize = 11; break;
                case 4:
                    _modelWork.FrontSize = 12; break;
            }
            switch (parameter.ToString())
            {
                case "expert":          
                    _testE=_modelWork.SelectExpert();  break;
                case "tehnadzor":
                    _testT = _modelWork.SelectTehnadzor();  break;
                case "workers":
                    if (_daysOrPeople)
                    {
                        _testPeople = _modelWork.SelectPeople();
                        _testDay = false;
                        AmountWorkers = _modelWork.GetAllWorkers();
                    }
                    else MessageService.ShowExclamation("Нажмите кнопку [Начать работу] чтобы задать интервал рабочих или рабочих дней!");
                    break;
                case "days":
                    if (_daysOrPeople)
                    {
                        _testDay = _modelWork.SelectDays();
                        _testPeople = false;
                        AmountDays = _modelWork.GetAllWorkDays();
                    }
                    else MessageService.ShowExclamation("Нажмите кнопку [Начать работу] чтобы задать интервал рабочих или рабочих дней!");
                    break;
            }
        }
        public bool CanExecuteAddSelectCommand(object parameter)
        {
            if (parameter.ToString() == "expert"|| parameter.ToString() == "tehnadzor"|| parameter.ToString() == "days" || parameter.ToString() == "workers")
                return true;
            else return false;
        }

        public void GetAllErrorSm(string error)
        {
            TextErrorSm = error;
        }
        public void GetAllErrorGr(string error)
        {
            TextErrorGr = error;
        }

        RelayCommand _addStartWork;
        public ICommand startWork
        {
            get
            {
                if (_addStartWork == null)
                    _addStartWork = new RelayCommand(ExecuteAddStartCommand, CanExecuteAddStartCommand);
                return _addStartWork;
            }
        }
        private static int TakeColor(int index)
        {
            int colNum = 0;
            switch (index)
            {
                case 0: colNum = 3; break;
                case 1: colNum = 10; break;
                case 2: colNum = 25; break;
                case 3: colNum = 6; break;
                case 4: colNum = 46; break;
                case 5: colNum = 33; break;
                case 6: colNum = 53; break;
                case 7: colNum = 1; break;
            }
            return colNum;
        }
        public async void ExecuteAddStartCommand(object parameter)
        {
            switch (parameter.ToString())
            {
                case "startSmeta":       
                    if(_flag)
                    {
                        if (_modelWork.AdressSmeta != null && _modelWork.AdressAktKS != null && _modelWork.AdressSaveSmeta != null)
                        {
                            _flag = false;
                            if (_testE)
                            {
                                await Task.Factory.StartNew(() =>
                                {
                                    _modelWork.StartProcessE();
                                    _flag = true;
                                    _exitOperation = true;
                                });
                                if (_exitOperation)
                                {
                                    if (_modelWork.TextError.Length == 0)
                                    {
                                        MessageService.ShowMessage("Ведомость эксперта успешно сохранена");
                                    }
                                    else
                                    {
                                        GetAllErrorSm(_modelWork.TextError);
                                        MessageService.ShowError("Устраните все ошибки и попробуйте снова");
                                    }
                                }
                            }

                            else if (_testT)
                            {
                                await Task.Factory.StartNew(() =>
                                {
                                    _modelWork.StartProcessT();
                                    _flag = true;
                                    _exitOperation = true;
                                });
                                if (_exitOperation)
                                {
                                    if (_modelWork.TextError.Length == 0)
                                    {
                                        MessageService.ShowMessage("Ведомость технадзора успешно сохранена");
                                    }
                                    else
                                    {
                                        GetAllErrorSm(_modelWork.TextError);
                                        MessageService.ShowError("Устраните все ошибки и попробуйте снова");
                                    }
                                }
                            }
                            else
                            {
                                MessageService.ShowExclamation("Вы не выбрали режим!");
                                _flag = true;
                            }
                        }
                        else MessageService.ShowExclamation("Вы пытаетесь выполнить работу, но не выбрали папку со сметами, актами КС и для сохранения!");
                    }
                    else MessageService.ShowExclamation("Вы уже запустили процесс обработки. Подождите!");
                    break;
                    case "startGraph":
                    if (_flag)
                    {
                        if (_modelWork.AdressOneSmeta != null && _modelWork.AdressSaveGraph != null && _daysOrPeople==true)
                        {
                        if (IndexColors == -1) _colorGet = TakeColor(0);
                        else _colorGet = TakeColor(IndexColors);
                        DateTime H = DateStart;
                        _modelWork.GetInputValueData(H.Day, H.Month, H.Year);
                        _flag = false;
                       
                            if (_testDay)
                            {
                                await Task.Factory.StartNew(() =>
                                {
                                    if (IndexAmountDays == -1) _modelWork.InputDays(AmountDays[0]);
                                    else _modelWork.InputDays(AmountDays[IndexAmountDays]);
                                    _modelWork.StartGraphDays(_colorGet);
                                    _flag = true;
                                    _exitOperation = true;
                                });
                                if (_exitOperation)
                                {
                                    if (_modelWork.TextError == null)
                                    {
                                        MessageService.ShowMessage("График успешно сохранен");
                                    }
                                    else
                                    {
                                        GetAllErrorGr(_modelWork.TextError);
                                        MessageService.ShowError("Устраните все ошибки и попробуйте снова");
                                    }
                                }
                            }
                            else if (_testPeople)
                            {

                                await Task.Factory.StartNew(() =>
                                {
                                    if (IndexAmountWorker == -1) _modelWork.InputPeople(AmountWorkers[0]);
                                    else _modelWork.InputPeople(AmountWorkers[IndexAmountWorker]);
                                    _modelWork.StartGraphPeople(_colorGet);
                                    _flag = true;
                                    _exitOperation = true;
                                });
                                if (_exitOperation)
                                {
                                    if (_modelWork.TextError == null)
                                    {
                                        MessageService.ShowMessage("График успешно сохранен");
                                    }
                                    else
                                    {
                                        GetAllErrorGr(_modelWork.TextError);
                                        MessageService.ShowError($"Устраните все ошибки и попробуйте снова");
                                    }
                                }
                            }
                            else 
                            {
                                MessageService.ShowExclamation("Вы не выбрали количество человек или дней!");
                                _flag = true;
                            }
                        }
                       else  MessageService.ShowExclamation("Вы пытаетесь выполнить работу, но не выбрали cмету, и папку для сохранения! Или же вы не нажали кгопку [Начать работу]");

                    }
                    else MessageService.ShowExclamation("Вы уже запустили процесс обработки. Подождите!");
            break;
            }
        }
        public bool CanExecuteAddStartCommand(object parameter)
        {
            if (parameter.ToString() == "startSmeta" || parameter.ToString() == "startGraph")
                return true;
            else return false;
        }

        RelayCommand _addFirstStartGraph;
        public ICommand AddFirstStart
        {
            get
            {
                if (_addFirstStartGraph == null)
                    _addFirstStartGraph = new RelayCommand(ExecuteAddFirstStartCommand, CanExecuteAddFirstStartCommand);
                return _addFirstStartGraph;
            }
        }
        public void ExecuteAddFirstStartCommand(object parameter)
        {

            if (_flag)
            {
                if (_modelWork.AdressOneSmeta != null && _modelWork.AdressSaveGraph != null)
                {
                    _modelWork.StartChoice();
                    _daysOrPeople = true;
                }
                else MessageService.ShowExclamation("Вы пытаетесь выполнить работу, но не выбрали cмету, и папку для сохранения!");
            }
            else MessageService.ShowExclamation("Вы уже запустили процесс обработки. Подождите!");
        }
        public bool CanExecuteAddFirstStartCommand(object parameter)
        {
            if (parameter.ToString() == "FirstStartGraph")
                return true;
            else return false;
        }

     


        RelayCommand _addExit;
        public ICommand AddExit
        {
            get
            {
                if (_addExit == null)
                    _addExit = new RelayCommand(ExecuteAddExitCommand, CanExecuteAddExitCommand);
                return _addExit;
            }
        }

        public void ExecuteAddExitCommand(object parameter)
        {
                if (parameter.ToString()== "exitSmeta"|| parameter.ToString() == "exitGraph")
                {
                if (_flag) { _modelWork.ExitError();  Application.Current.MainWindow.Close(); }
                else MessageService.ShowExclamation("Вы не можете выйти, производится работа над файлами"); 
                }
        }

        public bool CanExecuteAddExitCommand(object parameter)
        {
            if (parameter.ToString() == "exitSmeta" || parameter.ToString() == "exitGraph")
                return true;
            else return false;
        }
    }
}
