using Microsoft.Win32;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExectWork
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private string _filePath;
        private string _fileEX;
        private IWorkbook _workBook;
        private List<string> _sheetList=new List<string> ();
        private int _beginNumber = 0;
        List<Tuple<ICell, int, bool>> _cellList = new List<Tuple<ICell, int, bool>>();
        private void Open_button_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.ShowDialog();
            if (!string.IsNullOrEmpty(openFile.FileName) && openFile.CheckFileExists)
            {
                if (GetFileName(openFile.FileName))
                {
                    try
                    {
                        FileStream file = new FileStream(openFile.FileName, FileMode.Open);
                        _workBook = WorkbookFactory.Create(file);
                        
                        for (int i = 0; i < _workBook.NumberOfSheets; i++)
                        {
                            _sheetList.Add(_workBook.GetSheetAt(i).SheetName);
                        }
                        Template_Combobox.ItemsSource = _sheetList;
                        Data_ComboBox.ItemsSource = _sheetList;
                        int i1=_sheetList.IndexOf("模板");
                        int i2=_sheetList.IndexOf("数据");
                        if (i1!=-1 && i2!=-1)
                        {
                            Template_Combobox.SelectionChanged -= Template_Combobox_SelectionChanged;
                            Data_ComboBox.SelectionChanged -= Data_ComboBox_SelectionChanged;
                            Template_Combobox.SelectedIndex = i1;
                            Data_ComboBox.SelectedIndex = i2;
                            Template_Combobox.SelectionChanged += Template_Combobox_SelectionChanged;
                            Data_ComboBox.SelectionChanged += Data_ComboBox_SelectionChanged;
                        }
                        else
                        {
                            Template_Combobox.SelectionChanged -= Template_Combobox_SelectionChanged;
                            Data_ComboBox.SelectionChanged -= Data_ComboBox_SelectionChanged;
                            Template_Combobox.SelectedIndex = 0;
                            Data_ComboBox.SelectedIndex = 1;
                            Template_Combobox.SelectionChanged += Template_Combobox_SelectionChanged;
                            Data_ComboBox.SelectionChanged += Data_ComboBox_SelectionChanged;
                        }
                    }
                    catch (Exception ex)
                    {

                    }

                }
            }
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Analysis();
            WriteRows();

        }

        private void Analysis()
        {
            ISheet sheet1 = _workBook.GetSheet(Template_Combobox.SelectedItem as string);
            var dataList = sheet1.GetRowEnumerator();
            while (dataList.MoveNext())
            {
                IRow row = dataList.Current as IRow;
                if (row != null)
                {
                    foreach (ICell cell in row)
                    {
                        Tuple<ICell, int, bool> tuple;
                        if (AnalysisCell(cell, out tuple))
                        {
                            _cellList.Add(tuple);
                            //out_textBox.Text += tuple.Item1.Address.Row.ToString()+"::" + tuple.Item1.Address.Column.ToString()+"->" + tuple.Item2.ToString()+"&" + tuple.Item3.ToString() + "\n";
                        }
                    }
                }
            }
            out_textBox.Text += "分析完成：" + _cellList.Count.ToString() + "项。\n";
        }
        private bool AnalysisCell(ICell cell, out Tuple<ICell, int, bool> tuple)
        {
            bool rest = false;
            tuple = new Tuple<ICell, int, bool>(null, 0, false);
            if (cell != null)
            {                
                if (cell != null)
                {
                    switch (cell.CellType)
                    {
                        case CellType.Numeric:
                            if(!DateUtil.IsCellDateFormatted(cell))
                            {
                                double d = cell.NumericCellValue;
                                if ((d - Convert.ToInt32(d)) == 0)
                                {                                        
                                    tuple = new Tuple<ICell, int, bool>(cell, (int)(d- _beginNumber), false);                                        
                                    rest = true;                                    
                                }
                            }                            
                            break;
                        case CellType.String:
                            string st = cell.StringCellValue;
                            if (int.TryParse(st, out int value))
                            {
                                tuple = new Tuple<ICell, int, bool>(cell, value, false);
                                rest = true;
                            }
                            else
                            {
                                string[] vs = st.Split('&');
                                if (vs.Count() >= 3)
                                {
                                    tuple = new Tuple<ICell, int, bool>(cell, 0, true);
                                    rest = true;
                                }
                            }
                            break;
                        default:
                            rest = false;
                            break;
                    }
                }
            }
            return rest;
        }

        private void WriteRows()
        {
            ISheet sheet2 = _workBook.GetSheet(Data_ComboBox.SelectedItem as string);
            int i;
            for (i = 1; i <= sheet2.LastRowNum; i++)
            {
                IRow row = sheet2.GetRow(i);                
                foreach (Tuple<ICell, int, bool> t in _cellList)
                {
                    if (!t.Item3) CopyCell(row.GetCell(t.Item2), t.Item1);
                    else
                    {
                        string[] vs = t.Item1.StringCellValue.Split('&');
                        string st = vs[0];
                        for (int j = 1; j < vs.Count(); j++)
                        {
                            if (int.TryParse(vs[j], out int value))
                            {
                                ICell cell = row.GetCell(value);
                                st += GetCellString(cell);
                                j++;
                                st += vs[j];
                            }
                        }
                        t.Item1.SetCellValue(st);
                    }
                }
                string filename = _filePath + GetCellString(row.Cells[3]) + "." + _fileEX;
                FileStream file = new FileStream(filename, FileMode.Create);
                if (_workBook.NumberOfSheets == 2) _workBook.RemoveSheetAt(_workBook.GetSheetIndex(Data_ComboBox.SelectedItem as string));
                _workBook.Write(file);
                file.Close();               
                out_textBox.Text += GetCellString(row.Cells[3]) + "->完成" + "\n";                
            }
            out_textBox.Text += "共完成" + (i - 1).ToString() + "项目";
            MessageBox.Show("共完成" + (i - 1).ToString() + "项目");
        }

        private void CopyCell(ICell sourceCell,ICell targetCell)
        {
            if(sourceCell !=null)
            {
                switch (sourceCell.CellType )
                {
                    case CellType.Blank:                        
                        targetCell.SetCellValue("");
                        break;
                    case CellType.Boolean:
                        bool b = sourceCell.BooleanCellValue;
                        targetCell.SetCellValue(b);
                        targetCell.CellStyle.DataFormat = sourceCell.CellStyle.DataFormat;
                        break;
                    case CellType.Error:                        
                        targetCell.SetCellValue("");
                        break;
                    case CellType.Formula:                       
                        switch (sourceCell.CachedFormulaResultType)
                        {
                            case CellType.Numeric:
                                double d1 = sourceCell.NumericCellValue;
                                targetCell.SetCellType(CellType.Numeric);
                                targetCell.SetCellValue(d1);
                                targetCell.CellStyle.DataFormat = sourceCell.CellStyle.DataFormat;
                                break;
                            case CellType.String:
                                string st1 = sourceCell.StringCellValue;
                                targetCell.SetCellType(CellType.String);
                                targetCell.SetCellValue(st1);
                                targetCell.CellStyle.DataFormat = sourceCell.CellStyle.DataFormat;
                                break;
                            default:
                                out_textBox.Text += "CopyCell CellType.Formula default" + sourceCell.CachedFormulaResultType.ToString() + "\n";
                                break;
                        }
                        break;
                    case CellType.Numeric:
                        double d = sourceCell.NumericCellValue;
                        targetCell.SetCellType(CellType.Numeric);
                        targetCell.SetCellValue(d);                        
                        targetCell.CellStyle.DataFormat = sourceCell.CellStyle.DataFormat;                           
                        break;
                    case CellType.String:
                        string st = sourceCell.StringCellValue;
                        targetCell.SetCellType(CellType.String);
                        targetCell.SetCellValue(st);
                        targetCell.CellStyle.DataFormat = sourceCell.CellStyle.DataFormat;
                        break;
                    case CellType.Unknown:
                        targetCell.SetCellValue("");
                        break;
                }
            }


        }

        string GetCellString(ICell cell)
        {
            string st = "";
            if (cell == null) return st;
            switch (cell.CellType)
            {
                case CellType.Numeric:
                    if (!DateUtil.IsCellDateFormatted(cell)) st = cell.NumericCellValue.ToString();
                    else st = cell.DateCellValue.ToString(cell.CellStyle.GetDataFormatString());                   
                    break;
                case CellType.String:
                    st = cell.StringCellValue;
                    break;
                case CellType.Formula:
                    if (cell.CachedFormulaResultType == CellType.Numeric) st = cell.NumericCellValue.ToString(cell.CellStyle.GetDataFormatString());
                    break;
            }
            return st;
        }
        bool GetFileName(string st)
        {
            if (st != null)
            {
                string s, s1, s2;
                s1 = st.Split('\\').Last<string>();
                _filePath = st.Substring(0, st.Length - s1.Length);
                s = s1.Split('.').Last<string>();
                _fileEX = s;
                if (s == "xlsx" || s == "xls")
                {
                    return true;
                }

            }
            return false;
        }
        private void Data_ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Template_Combobox.SelectionChanged -= Template_Combobox_SelectionChanged;
            Data_ComboBox.SelectionChanged -= Data_ComboBox_SelectionChanged;
            int v1 = _sheetList.IndexOf(e.AddedItems[0] as string);
            int v2 = _sheetList.IndexOf(e.RemovedItems[0] as string);
            int v3 = Template_Combobox.SelectedIndex;
            if (v1 != v2 && v1 == v3) Template_Combobox.SelectedIndex = v2;
            Template_Combobox.SelectionChanged += Template_Combobox_SelectionChanged;
            Data_ComboBox.SelectionChanged += Data_ComboBox_SelectionChanged;
        }
        private void Template_Combobox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            Template_Combobox.SelectionChanged -= Template_Combobox_SelectionChanged;
            Data_ComboBox.SelectionChanged -= Data_ComboBox_SelectionChanged;
            int v1 = _sheetList.IndexOf(e.AddedItems[0] as string);
            int v2 = _sheetList.IndexOf(e.RemovedItems[0] as string);
            int v3 = Data_ComboBox.SelectedIndex;
            if (v1 != v2 && v1 == v3) Data_ComboBox.SelectedIndex = v2;                
            Template_Combobox.SelectionChanged += Template_Combobox_SelectionChanged;                
            Data_ComboBox.SelectionChanged += Data_ComboBox_SelectionChanged;           
        }
        
    }
}
