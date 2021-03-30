using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfNpoi
{
    public  class ExcelDocumentConverter
    {
        public static XSSFWorkbook convertWorkbookHSSFToXSSF(HSSFWorkbook source)
        {
            XSSFWorkbook retVal = new XSSFWorkbook();
            for (int i = 0; i < source.NumberOfSheets; i++)
            {
                XSSFSheet xssfSheet = (XSSFSheet)retVal.CreateSheet();
                HSSFSheet hssfsheet = (HSSFSheet)source.GetSheetAt(i);
                copySheets(hssfsheet, xssfSheet);
            }
            return retVal;
        }

        public static void copySheets(HSSFSheet source, XSSFSheet destination)
        {
            copySheets(source, destination, true);
        }

      
      
        public static void copySheets(HSSFSheet source, XSSFSheet destination, bool copyStyle)
        {
            int maxColumnNum = 0;
            Dictionary<int, HSSFCellStyle> styleMap = (copyStyle) ?new  Dictionary<int, HSSFCellStyle>() : null;
            for (int i = source.FirstRowNum ; i <= source.LastRowNum ; i++)
            {
                HSSFRow srcRow = (HSSFRow)source.GetRow(i);
                XSSFRow destRow =(XSSFRow) destination.CreateRow(i);
                if (srcRow != null)
                {
                    copyRow(source, destination, srcRow, destRow, styleMap);
                    if (srcRow.LastCellNum > maxColumnNum)
                    {
                        maxColumnNum = srcRow.LastCellNum;
                    }
                }
            }
            for (int i = 0; i <= maxColumnNum; i++)
            {                
                destination.SetColumnWidth(i, source.GetColumnWidth(i));
            }
        }

        /**
         * @param srcSheet
         *            the sheet to copy.
         * @param destSheet
         *            the sheet to create.
         * @param srcRow
         *            the row to copy.
         * @param destRow
         *            the row to create.
         * @param styleMap
         *            -
         */
        public static void copyRow(HSSFSheet srcSheet, XSSFSheet destSheet, HSSFRow srcRow, XSSFRow destRow,
                Dictionary<int, HSSFCellStyle> styleMap)
        {
            // manage a list of merged zone in order to not insert two times a
            // merged zone
            List<CellRangeAddress> mergedRegions = new List<CellRangeAddress>();
            destRow.Height=srcRow.Height ;
            // pour chaque row
            for (int j = srcRow.FirstCellNum; j <= srcRow.LastCellNum; j++)
            {
                HSSFCell oldCell = (HSSFCell)srcRow.GetCell (j); // ancienne cell
                XSSFCell newCell = (XSSFCell)destRow.GetCell(j); // new cell
                if (oldCell != null)
                {
                    if (newCell == null)
                    {
                        newCell = (XSSFCell)destRow.CreateCell (j);
                    }
                    // copy chaque cell
                    copyCell(oldCell, newCell, styleMap);
                    // copy les informations de fusion entre les cellules
                    // System.out.println("row num: " + srcRow.getRowNum() +
                    // " , col: " + (short)oldCell.getColumnIndex());
                    CellRangeAddress mergedRegion = getMergedRegion(srcSheet, srcRow.RowNum ,
                            (short)oldCell.ColumnIndex);

                    if (mergedRegion != null)
                    {
                        // System.out.println("Selected merged region: " +
                        // mergedRegion.toString());
                        CellRangeAddress newMergedRegion = new CellRangeAddress(mergedRegion.FirstRow ,
                                mergedRegion.LastRow, mergedRegion.FirstColumn, mergedRegion.LastColumn);
                        // System.out.println("New merged region: " +
                        // newMergedRegion.toString());

                        /*
                        CellRangeAddress wrapper = new CellRangeAddress(newMergedRegion);
                        CellRangeAddress ce=new CellRangeAddress ()
                        if (isNewMergedRegion(wrapper, mergedRegions))
                        {
                            mergedRegions.add(wrapper);
                            destSheet.addMergedRegion(wrapper.range);
                        }*/
                    }
                }
            }

        }

        /**
         * @param oldCell
         * @param newCell
         * @param styleMap
         */
        public static void copyCell(HSSFCell oldCell, XSSFCell newCell, Dictionary<int, HSSFCellStyle> styleMap)
        {
            if (styleMap != null)
            {
                int stHashCode = oldCell.CellStyle.GetHashCode();
                HSSFCellStyle sourceCellStyle = styleMap[stHashCode];
                XSSFCellStyle destnCellStyle = (XSSFCellStyle)newCell.CellStyle;
                if (sourceCellStyle == null)
                {                    
                    sourceCellStyle = (HSSFCellStyle)oldCell.Sheet.Workbook.CreateCellStyle();
                }                
                destnCellStyle.CloneStyleFrom(oldCell.CellStyle);
                
                styleMap.Add(stHashCode, sourceCellStyle);                
                newCell.CellStyle = destnCellStyle;
            }
            switch (oldCell.CellType )
            {
                case CellType.String:                    
                    newCell.SetCellValue(oldCell.StringCellValue);
                    break;
                case CellType.Numeric:
                    newCell.SetCellValue(oldCell.NumericCellValue);
                    break;
                case CellType.Blank:
                    newCell.SetCellValue(string.Empty); //!!!
                    break;
                case CellType.Boolean:
                    newCell.SetCellValue(oldCell.BooleanCellValue);
                    break;
                case CellType.Error:
                    newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                    break;
                case CellType.Formula:
                    newCell.SetCellFormula(oldCell.CellFormula);
                    break;
                default:
                    break;
            }

        }

        /**
         * Récupère les informations de fusion des cellules dans la sheet source
         * pour les appliquer à la sheet destination... Récupère toutes les zones
         * merged dans la sheet source et regarde pour chacune d'elle si elle se
         * trouve dans la current row que nous traitons. Si oui, retourne l'objet
         * CellRangeAddress.
         * 
         * @param sheet
         *            the sheet containing the data.
         * @param rowNum
         *            the num of the row to copy.
         * @param cellNum
         *            the num of the cell to copy.
         * @return the CellRangeAddress created.
         */
        public static CellRangeAddress getMergedRegion(HSSFSheet sheet, int rowNum, short cellNum)
        {
            
            for (int i = 0; i < sheet.NumMergedRegions; i++)
            {
                CellRangeAddress merged = sheet.MergedRegions[i];
                if (merged.IsInRange(rowNum, cellNum))
                {
                    return merged;
                }
                
            }
            
            return null;
        }

    
        private static bool isNewMergedRegion(CellRangeAddress newMergedRegion, CellRangeAddress[] mergedRegions)
        {
            return !mergedRegions.Contains(newMergedRegion);           
        }
    }
}
