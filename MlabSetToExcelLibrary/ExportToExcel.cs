using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MlabSetToExcelLibrary
{
    public static class ExportToExcel
    {
        private static string filename { get; set; }

        private static Excel.Application CreateExcelObj()
        {
            object obj;
            obj = null;
            try
            {
                //Создаём приложение.
                Excel.Application objExcel = new Excel.Application();
                obj = objExcel;

            }
            catch (Exception ex)
            {
                throw new Exception("Ошибка создания экземпляра MS Excel");
            }
            return (obj as Excel.Application);
        }

        private static void FormatSheetForSet2(Excel.Worksheet sheet, SetViewModel obj)
        {
            // formatting All sheet
            sheet.PageSetup.PrintGridlines = false;
            sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            sheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
            sheet.PageSetup.RightFooter = "Дата: &DD Стр &PP из &NN";
            sheet.PageSetup.RightHeader = "Исследование " + obj.Set.First().Project + ", сет № " + obj.Set.First().Set + " - " + obj.Set.First().TestMethod +
                                          " - " + obj.Set.First().AB;
            sheet.PageSetup.Zoom = false;
            sheet.PageSetup.LeftHeader = "НИИ Антимикробной химиотерапии";
            sheet.PageSetup.TopMargin = 50;
            sheet.PageSetup.BottomMargin = 50;
            sheet.PageSetup.HeaderMargin = 20;
            sheet.PageSetup.FooterMargin = 20;
            sheet.PageSetup.RightMargin = 10;
            sheet.PageSetup.LeftMargin = 50;
            sheet.PageSetup.Order = Excel.XlOrder.xlOverThenDown;

            // Formatting Set Number
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 3]].Merge();
            FormatHeaderText1(sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 3]]);
            // Formatting Footer cell
            sheet.Range[sheet.Cells[3 + obj.Set.First().MOList.Count + obj.Set.First().ControlMOList.Count, 1], sheet.Cells[3 + obj.Set.First().MOList.Count + obj.Set.First().ControlMOList.Count, 3]].Merge();
            FormatHeaderText1(sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 3]]);

            //Formatting table with MO
            FormatTableCells2(sheet.Range[sheet.Cells[1, 1], sheet.Cells[3 + obj.Set.First().MOList.Count + obj.Set.First().ControlMOList.Count, obj.Set.Count + 3]]);

            //Formatting AB Header
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, obj.Set.Count() + 3]].RowHeight = 18;
            sheet.Range[sheet.Cells[2, 1], sheet.Cells[2, obj.Set.Count() + 3]].RowHeight = 80;
            sheet.Range[sheet.Cells[2, 4], sheet.Cells[2, obj.Set.Count() + 3]].Orientation = 90;

            sheet.Range[sheet.Cells[obj.Set.First().MOList.Count + obj.Set.First().ControlMOList.Count + 3
              , 4], sheet.Cells[obj.Set.First().MOList.Count + obj.Set.First().ControlMOList.Count + 3, obj.Set.Count() + 3]].RowHeight = 80;
            sheet.Range[sheet.Cells[obj.Set.First().MOList.Count + obj.Set.First().ControlMOList.Count + 3
                , 4], sheet.Cells[obj.Set.First().MOList.Count + obj.Set.First().ControlMOList.Count + 3, obj.Set.Count() + 3]].Orientation = 90;

            sheet.Range[sheet.Cells[2, 1], sheet.Cells[obj.Set.First().MOList.Count + obj.Set.First().ControlMOList.Count + 3, 1]].ColumnWidth = 5;
            sheet.Range[sheet.Cells[2, 2], sheet.Cells[obj.Set.First().MOList.Count + obj.Set.First().ControlMOList.Count + 3, 2]].ColumnWidth = 9;
            sheet.Range[sheet.Cells[2, 3], sheet.Cells[obj.Set.First().MOList.Count + obj.Set.First().ControlMOList.Count + 3, 3]].ColumnWidth = 15;

            FormatHeaderControlMOText2(sheet.Range[sheet.Cells[obj.Set.First().MOList.Count + 3, 1], sheet.Cells[obj.Set.First().MOList.Count + obj.Set.First().ControlMOList.Count + 3, 3 + obj.Set.Count]]);




            sheet.Cells[obj.Set.First().MOList.Count + obj.Set.First().ControlMOList.Count + 7, 2] = "Проверил:";
        }
        private static void FormatSheetForSet1(Excel.Worksheet sheet, SetItem obj)
        {

            // formatting All sheet
            sheet.PageSetup.PrintGridlines = false;
            ;
            sheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            sheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
            sheet.PageSetup.RightFooter = "Дата: &DD Стр &PP из &NN";
            sheet.PageSetup.RightHeader = "Исследование " + obj.Project + ", сет № " + obj.Set + " - " + obj.TestMethod +
                                          " - " + obj.AB;
            sheet.PageSetup.Zoom = false;
            sheet.PageSetup.LeftHeader = "НИИ Антимикробной химиотерапии";
            sheet.PageSetup.TopMargin = 50;
            sheet.PageSetup.BottomMargin = 50;
            sheet.PageSetup.HeaderMargin = 20;
            sheet.PageSetup.FooterMargin = 20;
            sheet.PageSetup.RightMargin = 10;
            sheet.PageSetup.LeftMargin = 50;
            sheet.PageSetup.Order = Excel.XlOrder.xlOverThenDown;

            //// Foramatting test method
            //sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, obj.MICList.Count + 5]].Merge();
            //FormatHeaderText1(sheet.Range[sheet.Cells[1, 1], sheet.Cells[1, 3]]);

            //// Formatting Set Number
            //sheet.Range[sheet.Cells[3, 1], sheet.Cells[3, 4]].Merge();
            //FormatHeaderText1(sheet.Range[sheet.Cells[3, 1], sheet.Cells[3, 1]]);

            //// Formatting Set Number
            //sheet.Range[sheet.Cells[3, 5], sheet.Cells[3, obj.MICList.Count + 5]].Merge();
            //FormatHeaderText1(sheet.Range[sheet.Cells[3, 5], sheet.Cells[3, obj.MICList.Count + 5]]);

            //Formatting table with MO
            FormatTableCells1(sheet.Range[sheet.Cells[1, 1], sheet.Cells[1 + obj.MOList.Count, obj.MICList.Count + 5]]);
            sheet.PageSetup.Zoom = false;
            sheet.PageSetup.FitToPagesWide = 1;
         //   sheet.PageSetup.FitToPagesTall = 0;
            //Formatting Control MO Header
            sheet.Range[sheet.Cells[2 + obj.MOList.Count, 1], sheet.Cells[2 + obj.MOList.Count, obj.MICList.Count + 5]].Merge();
            sheet.Range[sheet.Cells[2 + obj.MOList.Count, 1], sheet.Cells[2 + obj.MOList.Count, obj.MICList.Count + 5]].RowHeight = 15;
            FormatHeaderControlMOText1(
                sheet.Range[
                    sheet.Cells[2 + obj.MOList.Count, 1], sheet.Cells[2 + obj.MOList.Count, obj.MICList.Count + 5]]);

          

            // Formatting table with control MO
            FormatTableCells1(sheet.Range[sheet.Cells[1 + obj.MOList.Count + 1, 1], sheet.Cells[1 + obj.MOList.Count + 1 + obj.ControlMOList.Count, obj.MICList.Count + 5]]);
            FormatHeaderControlMOText1(sheet.Range[
                    sheet.Cells[3 + obj.MOList.Count, 2], sheet.Cells[1 + obj.MOList.Count + obj.ControlMOList.Count, 4]]);
            //Formatting Top Row
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[5, obj.MICList.Count + 5]].ColumnWidth = 6;
            //Formatting Left Columns
            sheet.Range[sheet.Cells[1, 1], sheet.Cells[5 + obj.MOList.Count, 1]].ColumnWidth = 6;
            sheet.Range[sheet.Cells[1, 2], sheet.Cells[5 + obj.MOList.Count, 2]].ColumnWidth = 8;
            sheet.Range[sheet.Cells[1, 3], sheet.Cells[5 + obj.MOList.Count, 3]].ColumnWidth = 8;
            sheet.Range[sheet.Cells[1, 4], sheet.Cells[5 + obj.MOList.Count, 4]].ColumnWidth = 14;
            //Formatting Right Columns
            sheet.Range[sheet.Cells[1, obj.MICList.Count + 5], sheet.Cells[1 + obj.MOList.Count, obj.MICList.Count + 5]].ColumnWidth = 8;


            sheet.Cells[obj.MOList.Count + obj.ControlMOList.Count + 3, 2] = "Проверил:";

            // Разбиваем на две части
            if (obj.MOList.Count > 48)
            {
                sheet.ResetAllPageBreaks();
               // sheet.DisplayPageBreaks = true;
                sheet.HPageBreaks.Add(sheet.Cells[50, 1]);
            }
        }

        private static void FormatHeaderText1(Excel.Range range)
        {
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.RowHeight = 18;
            range.Font.Size = 10;
            range.Font.Bold = true;
        }

        private static void FormatHeaderControlMOText1(Excel.Range range)
        {
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            range.Font.Size = 10;
            range.Font.Bold = true;
            range.Font.Italic = true;
            range.Interior.ColorIndex = 34;
        }

        private static void FormatHeaderControlMOText2(Excel.Range range)
        {
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

            range.Font.Size = 10;
            range.Font.Bold = true;
            range.Font.Italic = true;
            range.Interior.ColorIndex = 34;
        }

        private static void FormatTableCells1(Excel.Range range)
        {
            range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Font.Size = 10;
            range.Font.Bold = true;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.WrapText = true;
            range.RowHeight = 19;
            range.ColumnWidth = 5;
        }

        private static void FormatTableCells2(Excel.Range range)
        {
            range.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Font.Size = 10;
            range.Font.Bold = true;
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            range.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
            range.WrapText = true;
            range.RowHeight = 30;
            range.ColumnWidth = 10;

        }

        public static int OpenDocument(string filepath, bool? csv)
        {
            var ExcelApp = CreateExcelObj();
            try
            {
                if (csv == true)
                {
                    ExcelApp.Visible = true;
                    var wb = ExcelApp.Workbooks.Open(filepath, 0, true, Excel.XlFileFormat.xlCSV, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    return 1;
                }
                else
                {
                    ExcelApp.Visible = true;
                    var wb = ExcelApp.Workbooks.Open(filepath, 0, true, Type.Missing, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                    return 1;
                }
               
            }

            catch
            {
                releaseObject(ExcelApp);
                return 0;
            }
        }

        public static int OpenExcelDocument(string filepath)
        {
            var ExcelApp = CreateExcelObj();
            try
            {
                //Excel.XlFileFormat.xlCSV
                ExcelApp.Visible = true;
                var wb = ExcelApp.Workbooks.Open(filepath, 0, true, Type.Missing, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                return 1;
            }

        catch
            {
                releaseObject(ExcelApp);
                return 0;
            }
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        } 

        public static string GetExcelDocumentSet(SetViewModel obj, string filePath, int setType)
        {
            Excel.Application ExcelApp;
            Excel.Worksheet ExcelSheet;
            Excel.Workbook ExcelWorkbook;
            Excel.Workbooks ExcelWorkbooks;
            Excel.Range ExcelRange;
            int rowsCount;
            int columnsCount;
            dynamic data;

            ExcelApp = CreateExcelObj();
            ExcelWorkbooks = ExcelApp.Workbooks;
            ExcelApp.ScreenUpdating = false;
            ExcelApp.DisplayAlerts = false;
            ExcelWorkbook = ExcelWorkbooks.Add();

            try
            {
              

                if (String.IsNullOrEmpty(filePath))
                {
                    filePath = Directory.GetParent(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)).FullName;
                    //if (Environment.OSVersion.Version.Major >= 6)
                    //{
                    //    filePath = Directory.GetParent(filePath).FullName;
                    //}
                }
                filename = obj.Set.First().Project + " - Сет " + obj.Set.First().Set + " - " + obj.Set.First().TestMethod + ".xlsx";




                switch (setType)
                {
                    case 1:
                       
                        foreach (var itemSet in obj.Set)
                        {
                            ExcelSheet = ExcelWorkbook.Sheets.Add();
                           
                            rowsCount = itemSet.MOList.Count + 8 + itemSet.ControlMOList.Count + 1;
                            columnsCount = itemSet.MICList.Count + 5;

                            ExcelRange =
                                ExcelSheet.Range[ExcelSheet.Cells[1, 1], ExcelSheet.Cells[rowsCount, columnsCount]];
                            if (
                            itemSet.AB.Length > 30)
                            {

                                ExcelSheet.Name = itemSet.AB.Substring(0, 30).Replace("/","|").Replace("\\","|");
                            }
                            else
                            {
                                ExcelSheet.Name = itemSet.AB.Replace("/", "|").Replace("\\", "|");
                            }

                            data = PrepareListForSet1(itemSet);

                            ExcelRange.Value = data;
                            FormatSheetForSet1(ExcelSheet, itemSet);
                       
                            Marshal.ReleaseComObject(ExcelRange);
                            Marshal.ReleaseComObject(ExcelSheet);

                        }
                        break;
                    case 2:
                        ExcelSheet = ExcelWorkbook.Sheets.Add();
                        rowsCount = obj.Set.First().MOList.Count + obj.Set.First().ControlMOList.Count + 3;
                        columnsCount = obj.Set.Count + 3;

                        ExcelRange = ExcelSheet.Range[ExcelSheet.Cells[1, 1], ExcelSheet.Cells[rowsCount, columnsCount]];

                        ExcelSheet.Name = obj.Set.First().Project + " - Сет № " + obj.Set.First().Set;

                        data = PrepareListForSet2(obj);

                        ExcelRange.Value = data;
                        FormatSheetForSet2(ExcelSheet, obj);
                        Marshal.ReleaseComObject(ExcelRange);
                        Marshal.ReleaseComObject(ExcelSheet);
                        break;
                    default:
                        break;
                }
                ExcelWorkbook.SaveAs();
                ExcelWorkbook.SaveAs(filePath + "\\" + filename, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);


                //while (Marshal.ReleaseComObject(ExcelWorkbook) > 0)
                //{ }
                //while (Marshal.ReleaseComObject(ExcelWorkbooks) > 0)
                //{ }


                //ExcelApp.Quit();

                //while (Marshal.ReleaseComObject(ExcelApp) > 0)
                //{ }

                return filePath + "\\" + filename;

            }
            catch (Exception ex)
            {

                throw ex;
            }
            finally
            {
                while (Marshal.ReleaseComObject(ExcelWorkbook) > 0)
                { }
                while (Marshal.ReleaseComObject(ExcelWorkbooks) > 0)
                { }


                ExcelApp.Quit();

                while (Marshal.ReleaseComObject(ExcelApp) > 0)
                { }
                GC.Collect();

            }
        }



        private static object[,] PrepareListForSet1(SetItem obj)
        {
            var rowsCount = obj.MOList.Count + 8 + obj.ControlMOList.Count + 1;
            var columnsCount = obj.MICList.Count + 5;
            object[,] data = new object[rowsCount, columnsCount];

            //data[0, 0] = "Метод тестирования: " + obj.TestMethod;
            //data[2, 0] = "Сет № " + obj.Set;
            //data[2, 4] = obj.AB;

            data[0, 0] = "Ячейка";
            data[0, 1] = "№";
            data[0, 2] = "Муз. №.";
            data[0, 3] = "МО";

            for (int i = 0; i < obj.MICList.Count; i++)
            {
                data[0, 4 + i] = obj.MICList[i];
            }

            data[0, 4 + obj.MICList.Count] = "МПК";

            for (int i = 0; i < obj.MOList.Count; i++)
            {
                data[1 + i, 0] = obj.MOList[i].Cell;
                data[1 + i, 1] = obj.MOList[i].Number;
                data[1 + i, 2] = obj.MOList[i].MuseumNumber;
                data[1 + i, 3] = obj.MOList[i].MO;

            }

            data[1 + obj.MOList.Count, 0] = "Контрольн.МО";

            for (int i = 0; i < obj.ControlMOList.Count; i++)
            {
                data[2 + obj.MOList.Count + i, 0] = obj.ControlMOList[i].Cell;
                data[2 + obj.MOList.Count + i, 1] = obj.ControlMOList[i].Number;
                data[2 + obj.MOList.Count + i, 2] = obj.ControlMOList[i].MuseumNumber;
                data[2 + obj.MOList.Count + i, 3] = obj.ControlMOList[i].MO;
            }

            return data;
        }

        private static object[,] PrepareListForSet2(SetViewModel obj)
        {
            var rowsCount = obj.Set.First().MOList.Count + obj.Set.First().ControlMOList.Count + 3;
            var columnsCount = obj.Set.Count + 3;
            object[,] data = new object[rowsCount, columnsCount];

            data[0, 0] = "Сет № " + obj.Set.First().Set;
            data[1, 0] = "№";
            data[1, 1] = "Муз. №";
            data[1, 2] = "МО";


            for (int i = 0; i < obj.Set.Count; i++)
            {
                data[0, 3 + i] = obj.Set[i].AB;
                data[1, 3 + i] = obj.Set[i].MICList.First().ToString() + " - " + obj.Set[i].MICList.Last().ToString();
                data[obj.Set.First().MOList.Count + obj.Set.First().ControlMOList.Count + 2, 3 + i] = obj.Set.First().ControlMICList.First().ToString() + " - " +
                                                           obj.Set[i].ControlMICList.Last().ToString();
            }

            for (int i = 0; i < obj.Set.First().MOList.Count; i++)
            {
                data[i + 2, 0] = obj.Set.First().MOList[i].Number;
                data[i + 2, 1] = obj.Set.First().MOList[i].MuseumNumber;
                data[i + 2, 2] = obj.Set.First().MOList[i].MO;

            }
            for (int i = 0; i < obj.Set.First().ControlMOList.Count; i++)
            {
                data[i + 2 + obj.Set.First().MOList.Count, 0] = obj.Set.First().ControlMOList[i].Number;
                data[i + 2 + obj.Set.First().MOList.Count, 1] = obj.Set.First().ControlMOList[i].MuseumNumber;
                data[i + 2 + obj.Set.First().MOList.Count, 2] = obj.Set.First().ControlMOList[i].MO;

            }


            return data;
        }
    }
}
