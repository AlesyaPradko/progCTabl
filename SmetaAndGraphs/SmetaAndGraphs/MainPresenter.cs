using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelEditor.bl;

namespace SmetaAndGraphs
{
   public class MainPresenter
    {
        private readonly IForm1 _view;
        private readonly IFileManager _manager;
        private readonly IMessageService _service;
        private bool _testE;
        private bool _testT;
        private bool _testDay;
        private bool _testPeople;
        private int _front=0;
        private int _days = 0;
        private int _workers = 0;
        private int _dayStart = 0;
        private int _monthStart = 0;
        private int _yearStart = 0;
        private int _colorGet = 0;
        public MainPresenter(IForm1 view, IFileManager manager, IMessageService service)
        {
            _view = view;
            _manager = manager;
            _service = service;
            _view.FileStartClik += _view_FileStartClik;
            _view.SelectModE += _view_SelectModE;
            _view.SelectModT += _view_SelectModT;
            _view.ChangeFront += _view_ChangeFront;
            _view.SelectDays += _view_SelectDays;
            _view.SelectPeople += _view_SelectPeople;
            _view.GraphFirstStartClik += _view_GraphFirstStartClik;
            _view.GraphStartClik += _view_GraphStartClik;
            _view.ChangeData += _view_ChangeData;
            _view.ChangeDays += _view_ChangeDays;
            _view.ChangePeople += _view_ChangePeople;
            _view.SelectColor += _view_SelectColor;
            _view.ExitError += _view_ExitError;
     
        }

        private void _view_ExitError(object sender, EventArgs e)
        {
            _manager.ExitError();
        }

        private void _view_SelectColor(object sender, EventArgs e)
        {
            string color = _view.ColorGraph;
            _colorGet = TakeColor(color);
        }
        private static int TakeColor(string col)
        {
            int colNum=0;
            switch (col)
            {
                case "красный": colNum = 3; break;
                case "зеленый": colNum = 10; break;
                case "синий": colNum = 25; break;
                case "желтый": colNum = 6; break;
                case "оранжевый": colNum = 46; break;
                case "голубой": colNum = 33; break;
                case "коричневый": colNum = 53; break;
                case "черный": colNum = 1; break;
            }
            return colNum;
        }
        private void _view_GraphFirstStartClik(object sender, EventArgs e)
        {

                _manager.InitializationFormGr(_view.OneSmetaAdres, _view.WhereSaveGraph);
                _manager.StartChoice();       
                _view.MinAmountDays = _manager.MinDays;
                _view.MaxAmountDays = _manager.MaxDays;
                _view.MinAmountPeople = _manager.MinPeople;
                _view.MaxAmountPeople = _manager.MaxPeople;
        }

        private void _view_ChangePeople(object sender, EventArgs e)
        {
            _workers = _view.AmountPeople;
            _manager.InputPeople(_workers);
            
        }

        private void _view_ChangeDays(object sender, EventArgs e)
        {
            _days = _view.AmountDays;
            _manager.InputDays(_days);
            
        }

        private void _view_ChangeData(object sender, EventArgs e)
        {
            _dayStart = _view.Days;
            _monthStart = _view.Month;
            _yearStart = _view.Year;
            _manager.GetInputValueData(_dayStart, _monthStart, _yearStart);          
        }

        private void _view_GraphStartClik(object sender, EventArgs e)
        {
            if (_colorGet == 0) _colorGet = TakeColor(_view.ColorGraph);
            if (_testDay)
            {
                Task taskBut = Task.Factory.StartNew(() =>
                {
                    if (_days == 0) _manager.InputDays(_view.AmountDays);
                    if (_dayStart == 0) _manager.GetInputValueData(_view.Days, _view.Month, _view.Year);
                    _manager.StartGraphDays(_colorGet);
                    if (_manager.TextError == null)
                    {
                        _service.ShowMessage("График успешно сохранен");
                    }
                    else
                    {
                        string error = _manager.TextError;
                        _service.ShowError($"{_manager.TextError}\n Устраните все ошибки и попробуйте снова");
                    }
                });

            }
            else if (_testPeople)
            {
                Task taskBut = Task.Factory.StartNew(() =>
                {
                    if (_workers == 0) _manager.InputPeople(_view.AmountPeople);
                    if (_dayStart == 0) _manager.GetInputValueData(_view.Days, _view.Month, _view.Year);
                    _manager.StartGraphPeople(_colorGet);
                    if (_manager.TextError == null)
                    {
                        _service.ShowMessage("График успешно сохранен");
                    }
                    else
                    {
                        string error = _manager.TextError;
                        _service.ShowError($"{_manager.TextError}\nУстраните все ошибки и попробуйте снова");  
                    }
                });

            }
            else
            {
                _service.ShowExclamation("Вы не выбрали количество человек или дней!");
            }
            _view.Flag = true;
        }

        private void _view_SelectPeople(object sender, EventArgs e)
        {
            _testPeople =_manager.SelectPeople();
            _testDay = false;
        }

        private void _view_SelectDays(object sender, EventArgs e)
        {
            _testDay=_manager.SelectDays();
            _testPeople = false;
        }

        private void _view_ChangeFront(object sender, EventArgs e)
        {
           _front = _view.SizeFront;
           _manager.FrontSize = _front;
        }

        private void _view_SelectModT(object sender, EventArgs e)
        {
            _testE = false;
            _testT = false;
            if (_front == 0)
            { _manager.FrontSize = _view.SizeFront; }
            _manager.InitializationFormTE(_view.SmetaAdres,_view.SmetaAktKS,_view.WhereSave);           
            _testT = _manager.SelectTehnadzor();
        }

        private void _view_SelectModE(object sender, EventArgs e)
        {
            _testE = false;
            _testT = false;
            if (_front == 0)
            { _manager.FrontSize = _view.SizeFront; }
            _manager.InitializationFormTE(_view.SmetaAdres, _view.SmetaAktKS, _view.WhereSave);
            _testE =_manager.SelectExpert();
        }

        private void _view_FileStartClik(object sender, EventArgs e)
        {

            if (_testE)
            {
                Task taskBut = Task.Factory.StartNew(() =>
                {
                    _manager.StartProcessE();

                    if (_manager.TextError.Length == 0)
                    {
                        _service.ShowMessage("Ведомость эксперта успешно сохранена");
                    }
                    else
                    {
                        string error = _manager.TextError;
                        _service.ShowError($"{_manager.TextError}\nУстраните все ошибки и попробуйте снова");
                    }
                });
            }
            else if (_testT)
            {
                Task taskBut = Task.Factory.StartNew(() =>
                {
                    _manager.StartProcessT();
                    if (_manager.TextError.Length == 0)
                    {
                        _service.ShowMessage("Ведомость технадзора успешно сохранена");
                    }
                    else
                    {
                        _service.ShowError($"{_manager.TextError}\nУстраните все ошибки и попробуйте снова");
                    }
                });
            }
            else
            {
                _service.ShowExclamation("Вы не выбрали режим!");
            }
            _view.Flag = true;
        }
    }
}
