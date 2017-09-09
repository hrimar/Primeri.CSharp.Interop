using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InteropExcel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Excel
{
    public class IOWrite
    {
        private DataStruct _data;
		private InteropExcel.Application excel;

        public IOWrite(DataStruct data)
        {
            _data = data;
        }

        public bool ExportTable()
        {
            try
            {
                // Подготовка
                excel = new InteropExcel.ApplicationClass();

                if (excel == null) return false;

                excel.Visible = false;


                InteropExcel.Workbook workbook = excel.Workbooks.Add();
                if (workbook == null) return false;

                InteropExcel.Worksheet sheet = (InteropExcel.Worksheet) workbook.Worksheets[1];
                sheet.Name = "Таблица 1";

                // Попълване на таблицата

                int i = 1;

                AddRow(new DataRow("Първо име", "Фамилия", "Години"), i++, true, 50); i++;
                foreach (DataRow row in _data.table)
                {

                    AddRow(row, i++, false, -1) ;
                }

                i++; AddRow(new DataRow("Брой редове", "", _data.table.Count.ToString()), i++, true, -1); 

                // Запаметяване и затваряне
                workbook.SaveCopyAs(GetPath());

                excel.DisplayAlerts = false; // Изключваме всички съобщения на Excel

                workbook.Close();
                excel.Quit();

                // Освобождаване на памет от Excel
                if (workbook != null) Marshal.ReleaseComObject(workbook);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
                if (excel != null) Marshal.ReleaseComObject(excel);

                workbook = null;
                sheet = null;
                excel = null;

                GC.Collect();

                return true;
            }
            catch 
            {
            }
            return false;
        }

        public void AddRow(DataRow dataRow, int indexRow, bool isBold, int color)
        {
            try
            {
                InteropExcel.Range range;

                //Фоорматиране                
                range = excel.Range["A" + indexRow.ToString(), "C" + indexRow.ToString()];
                if(color >0)     range.Interior.ColorIndex = color; // -1
                if (isBold)      range.Font.Bold = isBold;

                //Въвеждане данни клетка по клетка
                range = excel.Range["A" + indexRow.ToString(), "A" + indexRow.ToString()];
                range.Value2 = dataRow.FirstName;

                range = excel.Range["B" + indexRow.ToString(), "B" + indexRow.ToString()];
                range.Value2 = dataRow.LastName;

                range = excel.Range["C" + indexRow.ToString(), "C" + indexRow.ToString()];
                range.Value2 = dataRow.Age;
            }
            catch 
            {
            }
        }

        public void RunFile()
        {
            try
            {
                System.Diagnostics.Process.Start(GetPath());
            }
            catch 
            {
            }
        }

        private string GetPath()
        {
            return System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Table1.xlsx");
        }
    }
}
