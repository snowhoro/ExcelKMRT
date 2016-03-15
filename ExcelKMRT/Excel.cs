using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using xls = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;
using System.Runtime.InteropServices;
namespace WindowsFormsApplication1
{  
    class Excel
    {
        xls.Application xlApp;
        xls.Workbooks xlWorkBooks;
        xls.Workbook xlWorkBook;
        xls.Sheets xlSheets;
        xls.Worksheet xlWorkSheet;

        string filepath;
        List<DateTime> monthDates;
        int monthNumber;
        string monthName;
        int yearNumber;
        int minValue;
        int maxValue;
        int lastDay;
        bool RunTotalKM;
        int mondayKM;
        bool exists;

        public Excel(string _filepath, int _monthNumber, string _monthName, int _yearNumber, 
            int _minValue, int _maxValue, bool _RunTotalKM, int _mondayKM)
        {
            filepath = _filepath;
            monthNumber = _monthNumber;
            monthName = _monthName;
            yearNumber = _yearNumber;
            minValue = _minValue;
            maxValue = _maxValue;
            RunTotalKM = _RunTotalKM;
            mondayKM = _mondayKM;
        }

        public void StartExcel()
        {
            xlApp = new xls.Application();
            if (isExcelInstalled())
                RunExcel();
        }
        private void RunExcel()
        {
            if (!CheckFile())
                return;

            if (!OpenWorkSheet())
                CreateWorkSheet();

            TitleBar();
            Dates();
            Formulas();
            if(RunTotalKM)
                TotalKM();
            ResumenTotal();
            AdditionalData();
            Formats();

            SaveFile();
        }
        public void CloseExcel()
        {
            if (xlWorkSheet != null) 
                Marshal.ReleaseComObject(xlWorkSheet);
            if (xlSheets != null)
                Marshal.ReleaseComObject(xlSheets);
            if(xlWorkBook != null)
                Marshal.ReleaseComObject(xlWorkBook);
            if(xlWorkBooks != null)
                Marshal.ReleaseComObject(xlWorkBooks);
            if(xlApp != null)
                Marshal.ReleaseComObject(xlApp);

            xlWorkSheet = null;
            xlSheets = null;
            xlWorkBook = null;
            xlWorkBooks = null;
            xlApp = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        private void SaveFile()
        {
            if (exists)
                xlWorkBook.Save();
            else
                xlWorkBook.SaveAs(filepath);
            xlWorkBook.Close(true,filepath,System.Reflection.Missing.Value);
            xlApp.Quit();

            //MessageBox.Show("Excel creado, se encuentra en " + filename);
            Process.Start(filepath);
        }
        private bool CheckFile()
        {
            xlWorkBooks = xlApp.Workbooks;

            bool isFileOpen;
            if (!isFileCreated())
            {
                exists = false;
                isFileOpen = CreateFile();
            }
            else
            {
                exists = true;
                isFileOpen = OpenFile();
            }
            return isFileOpen;
        }

        private bool OpenWorkSheet()
        {
            xlSheets = xlWorkBook.Worksheets;
            foreach (xls.Worksheet sheet in xlSheets)
            {
                if (sheet.Name == monthName)
                {
                    DialogResult dResult = MessageBox.Show("Sobreescribir la hoja de" + monthName + "?", "Sobreescribir", MessageBoxButtons.YesNo);
                    if (dResult == DialogResult.Yes)
                    {
                        xlApp.DisplayAlerts = false;
                        sheet.Delete();
                        xlApp.DisplayAlerts = true;
                        xlWorkSheet = xlSheets.Add();
                        xlWorkSheet.Name = monthName;
                        return true;
                    }
                    else
                        return false;
                }
            }
            return false;
        }
        private void CreateWorkSheet()
        {
            xlWorkSheet = xlSheets.Add();
            xlWorkSheet.Name = monthName;
        }

        private bool CreateFile()
        {
            try
            {
                xlWorkBook = xlWorkBooks.Add();
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return false;
            }
        }
        private bool OpenFile()
        {
            if(isFileLocked())
                return false;

            try
            {
                xlWorkBook = xlWorkBooks.Open(filepath);
                return true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
                return false;
            }
        }

        private bool isExcelInstalled()
        {
            if (xlApp == null)
            {
                MessageBox.Show("Excel no esta instalado");
                return false;
            }
            return true;
        }
        private bool isFileCreated()
        {
            if (File.Exists(filepath))
                return true;
            return false;
        }
        private bool isFileLocked()
        {
            try
            {
                FileStream fs = File.OpenWrite(filepath);
                fs.Close();
                return false;
            }
            catch (Exception) 
            {
                MessageBox.Show("El archivo esta bloqueado por otro programa");
                return true; 
            }
        }

        private void TitleBar()
        {
            xlWorkSheet.Cells[1, 1] = "DIA";
            xlWorkSheet.Columns[1].ColumnWidth = 10;
            xlWorkSheet.Cells[1, 2] = "SALIDA KM";
            xlWorkSheet.Columns[2].ColumnWidth = 10;
            xlWorkSheet.Cells[1, 3] = "VUELTA KM";
            xlWorkSheet.Columns[3].ColumnWidth = 10.60;
            xlWorkSheet.Cells[1, 4] = "TOTAL KM REALIZADOS";
            xlWorkSheet.Columns[4].ColumnWidth = 21;
            xlWorkSheet.Cells[1, 5] = "A CARGO RT";
            xlWorkSheet.Columns[5].ColumnWidth = 11;
            xlWorkSheet.Cells[1, 6] = "DESTINOS";
            xlWorkSheet.Columns[6].ColumnWidth = 55;
        }
        private void Dates()
        {
            monthDates = AllDatesInAMonth(monthNumber + 1, yearNumber);
            xls.Range formatRange;
            for (int i = 0; i < monthDates.Count; i++)
            {
                if (monthDates[i].DayOfWeek == DayOfWeek.Saturday || monthDates[i].DayOfWeek == DayOfWeek.Sunday)
                {
                    formatRange = xlWorkSheet.get_Range("a1", "f" + monthDates.Count);
                    for (int j = 1; j < 7; j++)
                    {
                        xls.Range cell = formatRange.Cells[i + 2, j];
                        cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    }
                }
                xlWorkSheet.Cells[i + 2, 1] = monthDates[i];
            }
            formatRange = xlWorkSheet.get_Range("a2", "a" + monthDates.Count + 1);
            formatRange.NumberFormat = "dd/mm/yyyy";
        }
        private List<DateTime> AllDatesInAMonth(int month, int year)
        {
            var firstOftargetMonth = new DateTime(year, month, 1);
            var firstOfNextMonth = firstOftargetMonth.AddMonths(1);

            var allDates = new List<DateTime>();

            for (DateTime date = firstOftargetMonth; date < firstOfNextMonth; date = date.AddDays(1))
            {
                allDates.Add(date);
            }
            return allDates;
        }
        private void Formulas()
        {
            for (int i = 0; i < monthDates.Count; i++)
            {
                if (monthDates[i].DayOfWeek == DayOfWeek.Saturday || monthDates[i].DayOfWeek == DayOfWeek.Sunday)
                {
                    xlWorkSheet.Cells[i + 2, 2] = "";
                    xlWorkSheet.Cells[i + 2, 3] = "";
                }
                else if (monthDates[i].DayOfWeek == DayOfWeek.Monday)
                {
                    xlWorkSheet.Cells[i + 2, 2].Formula = "=C" + (i - 1); // menos 3 sab doming y anterior
                    xlWorkSheet.Cells[i + 2, 3].Formula = "=SUM(B" + (i + 2) + "+D" + (i + 2) + ")";
                    xlWorkSheet.Cells[i + 2, 5].Formula = "=D" + (i + 2) + "*60/100";
                }
                else
                {
                    xlWorkSheet.Cells[i + 2, 2].Formula = "=C" + (i + 1); //menos 1 dia anterior
                    xlWorkSheet.Cells[i + 2, 3].Formula = "=SUM(B" + (i + 2) + "+D" + (i + 2) + ")";
                    xlWorkSheet.Cells[i + 2, 5].Formula = "=D" + (i + 2) + "*60/100";
                }
            }

            //forced - FALTA OBTENER
            xlWorkSheet.Cells[2, 2] = 1000;
        }
        private void TotalKM()
        {
            Random rnd = new Random();
            int lastNumber = rnd.Next(minValue, maxValue);
            int newNumber;

            for (int i = 0; i < monthDates.Count; i++)
            {
                if (monthDates[i].DayOfWeek == DayOfWeek.Saturday || monthDates[i].DayOfWeek == DayOfWeek.Sunday)
                    xlWorkSheet.Cells[i + 2, 4] = "";
                else if (monthDates[i].DayOfWeek == DayOfWeek.Monday)
                {
                    xlWorkSheet.Cells[i + 2, 4] = mondayKM;
                    lastNumber = mondayKM;
                }
                else
                {
                    do
                    {
                        newNumber = rnd.Next(minValue, maxValue);
                    } while (newNumber == lastNumber);
                    lastNumber = newNumber;
                    xlWorkSheet.Cells[i + 2, 4] = newNumber;
                }

            }
        }
        private void ResumenTotal()
        {
            lastDay = monthDates.Count + 2;
            xlWorkSheet.Cells[lastDay, 3] = "TOTAL KM";
            xlWorkSheet.Cells[lastDay, 3].Font.Bold = true;
            xlWorkSheet.Cells[lastDay, 3].HorizontalAlignment = xls.XlHAlign.xlHAlignRight;

            xlWorkSheet.Cells[lastDay, 4].Formula = "=SUM(D2:D" + (lastDay - 1) + ")";
            xlWorkSheet.Cells[lastDay, 5].Formula = "=SUM(E2:E" + (lastDay - 1) + ")";
            xlWorkSheet.Cells[lastDay + 1, 4] = "TOTAL";
            xlWorkSheet.Cells[lastDay + 1, 4].Font.Bold = true;
            xlWorkSheet.Cells[lastDay + 1, 4].HorizontalAlignment = xls.XlHAlign.xlHAlignRight;

            xlWorkSheet.Cells[lastDay + 1, 5].Formula = "=E" + lastDay + "*2.96"; //cambio km pesos
        }
        private void Formats()
        {
            lastDay += 3;
            //TITLE FORMAT -------------------------------------
            xls.Range formatRange;
            formatRange = xlWorkSheet.get_Range("A4");
            formatRange.EntireRow.Font.Bold = true;
            xlWorkSheet.get_Range("A4", "F" + lastDay).HorizontalAlignment = xls.XlHAlign.xlHAlignCenter;

            formatRange = xlWorkSheet.UsedRange;
            for (int i = 1; i < 7; i++)
            {
                xls.Range cell = formatRange.Cells[4, i];
                xls.Borders border = cell.Borders;
                border.LineStyle = xls.XlLineStyle.xlContinuous;
                border.Weight = xls.XlBorderWeight.xlMedium;
            }
            //-------------------------------------------------
            // RESUMEN TOTAL ----------------------------------
            formatRange = xlWorkSheet.get_Range("a5", "e" + (lastDay - 1));
            formatRange.BorderAround(xls.XlLineStyle.xlContinuous,
            xls.XlBorderWeight.xlMedium, xls.XlColorIndex.xlColorIndexAutomatic,
            xls.XlColorIndex.xlColorIndexAutomatic);

            formatRange = xlWorkSheet.get_Range("f5", "f" + (lastDay - 1));
            formatRange.BorderAround(xls.XlLineStyle.xlContinuous,
            xls.XlBorderWeight.xlMedium, xls.XlColorIndex.xlColorIndexAutomatic,
            xls.XlColorIndex.xlColorIndexAutomatic);

            formatRange = xlWorkSheet.get_Range("a" + lastDay, "e" + (lastDay + 1));
            formatRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
            //-------------------------------------------------

        }
        private void AdditionalData()
        {
            for (int i = 0; i < 3; i++)
                xlWorkSheet.Rows[1].Insert();

            xlWorkSheet.Cells[1, 1] = "DATOS EMPRESA:  RADIATION ONCOLOGY SA           DATOS EMPLEADA:  GABRIELA S IBARRA           DATOS VEHICULO Peugeot 207";
            xlWorkSheet.get_Range("A1", "F1").Merge(false);
            xlWorkSheet.Cells[1, 1].Characters[17, 21].Font.Bold = true;
            xlWorkSheet.Cells[1, 1].Characters[66, 17].Font.Bold = true;
            xlWorkSheet.Cells[1, 1].Characters[109, 11].Font.Bold = true;

            xlWorkSheet.Cells[2, 1] = "                                      CRAMER 1180 CABA                                                                  18 389778                                                                    MHF 850";
            xlWorkSheet.get_Range("A2", "F2").Merge(false);
            xlWorkSheet.Cells[2, 1].Font.Bold = true;
            xlWorkSheet.Cells[2, 1].HorizontalAlignment = xls.XlHAlign.xlHAlignLeft;

            xlWorkSheet.Cells[3, 1] = "                                      30- 67817931-7                                                                            27-18389778-6";
            xlWorkSheet.get_Range("A3", "F3").Merge(false);
            xlWorkSheet.Cells[3, 1].Font.Bold = true;
            xlWorkSheet.Cells[3, 1].HorizontalAlignment = xls.XlHAlign.xlHAlignLeft;
        }
    }
}