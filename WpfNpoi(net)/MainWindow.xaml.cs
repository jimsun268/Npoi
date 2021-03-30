using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;

namespace WpfNpoi_net_
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            _openFile = new OpenFileDialog();
            _openFile.Filter = "XLS (*.xls)|*.xls|XLSX (*.xlsx)|*.xlsx";
        }

        private OpenFileDialog _openFile;
        private string _filePathAndName;
        private IWorkbook _workBook;

        public string FileExt
        {
            get
            {
                if (string.IsNullOrEmpty(_filePathAndName)) GetFile();
                return _filePathAndName.Split('.').Last();
            }
        }
        public string FileName
        {
            get
            {
                if (string.IsNullOrEmpty(_filePathAndName)) GetFile();
                return _filePathAndName.Split('\\').Last().Split('.').First();
            }
        }
        public string FilePath
        {
            get
            {
                if (string.IsNullOrEmpty(_filePathAndName)) GetFile();
                int indx = FileName.Length + FileExt.Length + 1;
                return _filePathAndName.Remove(_filePathAndName.Length - indx, indx);
            }
        }
        private void button_open_Click(object sender, RoutedEventArgs e)
        {
            GetFile();
            GetWorkBook();
        }

        private void button_Begin_Click(object sender, RoutedEventArgs e)
        {
            BeginCheck();
            SaveWorkBook();
        }

        private void GetFile()
        {

            if (_openFile.ShowDialog() == true)
            {
                _filePathAndName = _openFile.FileName;
                textBox_out.Text += $"已经选择{_filePathAndName }文件！ \n";
            }

        }
        private void GetWorkBook()
        {
            try
            {
                Stream filestream = _openFile.OpenFile();
                _workBook = WorkbookFactory.Create(filestream);
                if (_workBook != null)
                {
                    List<string> vs = new List<string>();
                    for (int i = 0; i < _workBook.NumberOfSheets; i++)
                    {
                        vs.Add(_workBook.GetSheetAt(i).SheetName);
                    }
                    Combobox_SheetName.ItemsSource = vs;
                    Combobox_SheetName.SelectedIndex = 0;
                    textBox_out.Text += $"成功打开{FileName }文件！\n";
                }

            }
            catch (Exception ex)
            {

                textBox_out.Text += ex.Message + "\n";
            }
        }
        private void BeginCheck()
        {
            textBox_out.Text += $"开始分析{FileName }文件！\n";
            ISheet sheet = _workBook.GetSheet(Combobox_SheetName.SelectedItem as string);
            ChectSheet(sheet);
        }
        private void ChectSheet(ISheet xsheet)
        {
            var datalist = xsheet.GetEnumerator();
            while (datalist.MoveNext())
            {
                IRow xrow = datalist.Current as IRow;
                if (xrow != null)
                {

                    foreach (ICell cell in xrow)
                    {
                        Dictionary<short, List<string>> _cellList;

                        if (FileExt == "xlsx") _cellList = AnalysisCellString(cell as XSSFCell);
                        else _cellList = AnalysisCellString(cell as HSSFCell);
                        foreach (short s in _cellList.Keys)
                        {
                            string st = "OutSheetColoris" + s.ToString();
                            ISheet sheet2;
                            if (_workBook.GetSheetIndex(st) < 0) sheet2 = _workBook.CreateSheet(st);
                            else sheet2 = _workBook.GetSheetAt(_workBook.GetSheetIndex(st));
                            List<string> vs = _cellList[s];
                            string st2 = string.Empty;
                            foreach (string t in vs)
                            {
                                st2 += t;
                            }
                            IRow r1;
                            if (sheet2.GetRow(cell.RowIndex) == null) r1 = sheet2.CreateRow(cell.RowIndex);
                            else r1 = sheet2.GetRow(cell.RowIndex);
                            ICell cell1 = r1.CreateCell(cell.ColumnIndex);
                            cell1.SetCellValue(st2);
                        }
                    }
                }
            }
        }
        private Dictionary<short, List<string>> AnalysisCellString(HSSFCell cell)
        {
            Dictionary<short, List<string>> _cellList = new Dictionary<short, List<string>>();
            HSSFRichTextString rich = (HSSFRichTextString)cell.RichStringCellValue;

            int formattingRuns = cell.RichStringCellValue.NumFormattingRuns;
            if (formattingRuns == 0) return _cellList;
            IFont font2 = _workBook.GetFontAt(cell.CellStyle.FontIndex);
            string st2 = rich.String.Substring(0, rich.GetIndexOfFormattingRun(0));
            if (_cellList.Keys.Where(t => t == font2.Color).Count() == 0) _cellList.Add(font2.Color, new List<string>() { st2 });
            else _cellList[font2.Color].Add(st2);
            for (int i = 0; i < formattingRuns; i++)
            {
                int startIdx = rich.GetIndexOfFormattingRun(i);
                int length;
                if (i == formattingRuns - 1) length = rich.Length - startIdx;
                else length = rich.GetIndexOfFormattingRun(i + 1) - startIdx;
                string st = rich.String.Substring(startIdx, length);

                short fontIndex = rich.GetFontOfFormattingRun(i);
                IFont font = _workBook.GetFontAt(fontIndex);

                if (_cellList.Keys.Where(t => t == font.Color).Count() == 0) _cellList.Add(font.Color, new List<string>() { st });
                else _cellList[font.Color].Add(st);
            }
            return _cellList;
        }
        private Dictionary<short, List<string>> AnalysisCellString(XSSFCell cell)
        {
            Dictionary<short, List<string>> _cellList = new Dictionary<short, List<string>>();
            XSSFRichTextString rich = (XSSFRichTextString)cell.RichStringCellValue;
            int formattingRuns = cell.RichStringCellValue.NumFormattingRuns;
            for (int i = 0; i < formattingRuns; i++)
            {
                int startIdx = rich.GetIndexOfFormattingRun(i);
                int length = rich.GetLengthOfFormattingRun(i);
                string st = rich.String.Substring(startIdx, length);
                IFont font;
                if (i == 0)
                {
                    short fontIndex = cell.CellStyle.FontIndex;
                    font = _workBook.GetFontAt(fontIndex);
                }
                else
                {
                    font = rich.GetFontOfFormattingRun(i);
                }
                if (_cellList.Keys.Where(t => t == font.Color).Count() == 0) _cellList.Add(font.Color, new List<string>() { st });
                else _cellList[font.Color].Add(st);
            }
            return _cellList;
        }
        private void SaveWorkBook()
        {
            string fp = FilePath + FileName + "out." + FileExt;
            FileStream file = new FileStream(fp, FileMode.Create);
            _workBook.Write(file);
            file.Close();
            textBox_out.Text += $"成功保存文件{fp }！\n";

        }
    }
}
