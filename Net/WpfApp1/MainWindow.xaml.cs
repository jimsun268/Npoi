using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace WpfApp1
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {

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
        public MainWindow()
        {
            InitializeComponent();
            _openFile = new OpenFileDialog();
            _openFile.Filter = "XLS (*.xls)|*.xls|XLSX (*.xlsx)|*.xlsx";
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
                    ICell cell1 = xrow.GetCell(0);
                    IFont font = cell1.CellStyle.GetFont(_workBook);
                    foreach (ICell cell in xrow)
                    {
                        CheckCell(cell);
                    }
                }
            }
        }

        private void CheckCell(ICell cell)
        {

            if (cell != null)
            {
                switch (cell.CellType)
                {
                    case CellType.Blank:
                        WriteCell(cell);
                        break;
                    case CellType.Boolean:
                        WriteCell(cell);
                        break;
                    case CellType.Error:
                        WriteCell(cell);
                        break;
                    case CellType.Formula:
                        WriteCell(cell);
                        break;
                    case CellType.Numeric:
                        WriteCell(cell);
                        break;
                    case CellType.String:
                        if (FileExt == "xlsx") AnalysisCellString(cell as XSSFCell);
                        else AnalysisCellString(cell as HSSFCell);
                        break;
                    case CellType.Unknown:
                        
                        break;
                }
            }
           
        }

        private void AnalysisCellString(XSSFCell cell)
        {           
            XSSFRichTextString rich = (XSSFRichTextString)cell.RichStringCellValue;
            string st2 = rich.String ;
            int formattingRuns = cell.RichStringCellValue.NumFormattingRuns;
            if (formattingRuns > 0) st2 = string.Empty;
            XSSFFont font=(XSSFFont)_workBook.CreateFont ();
            for (int i = 0; i < formattingRuns; i++)
            {
                int startIdx = rich.GetIndexOfFormattingRun(i);
                int length = rich.GetLengthOfFormattingRun(i);
                string st = rich.String.Substring(startIdx, length);
                
                if (i == 0)
                {
                    short fontIndex = cell.CellStyle.FontIndex;
                    font = (XSSFFont)_workBook.GetFontAt(fontIndex);                    
                }
                else
                {
                    font = (XSSFFont)rich.GetFontOfFormattingRun(i);
                }
                
                if (font.Color != IndexedColors.White.Index && font.Color != 0)
                {
                    
                    st2 += st;
                }
               
            }
            font.Color = IndexedColors.Black.Index;
            XSSFRichTextString rich2 = new XSSFRichTextString();
            rich2.Append(st2,font );
            cell.SetCellValue(rich2);
        }

        private void AnalysisCellString(HSSFCell cell)
        {
            string st3 = string.Empty;
            HSSFRichTextString rich = (HSSFRichTextString)cell.RichStringCellValue;

            int formattingRuns = cell.RichStringCellValue.NumFormattingRuns;
            if (formattingRuns == 0) return ;
            IFont font2 = _workBook.GetFontAt(cell.CellStyle.FontIndex);
            string st2 = rich.String.Substring(0, rich.GetIndexOfFormattingRun(0));
            if (font2.Color != IndexedColors.White.Index && font2.Color != 0)
            {
                st3 += st2;
            }
            for (int i = 0; i < formattingRuns; i++)
            {
                int startIdx = rich.GetIndexOfFormattingRun(i);
                int length;
                if (i == formattingRuns - 1) length = rich.Length - startIdx;
                else length = rich.GetIndexOfFormattingRun(i + 1) - startIdx;
                string st = rich.String.Substring(startIdx, length);

                short fontIndex = rich.GetFontOfFormattingRun(i);
                IFont font = _workBook.GetFontAt(fontIndex);
                if (font.Color != IndexedColors.White.Index && font.Color != 0)
                {
                    font2 = font;
                    st3 += st;
                }

            }
            HSSFRichTextString rich2 = new HSSFRichTextString(st3);
            rich2.ApplyFont(font2);
            cell.SetCellValue(rich2);
        }
        

        private void WriteCell(ICell cell)
        {
            var v=cell.CellStyle.GetFont(_workBook );
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
