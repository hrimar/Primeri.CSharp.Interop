using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel
{
    class Program
    {
        static void Main()
        {
            DataStruct data = new DataStruct();

            IOWrite write = new IOWrite(data);

            // Набиранена данни в основната таблица
            data.AddRow("Христо", "Христов", "42");
            data.AddRow("Ясен", "Петров", "43");

            // Проверка на таблицата
            data.PrintTable();

            write.ExportTable();
            write.RunFile();
        }
    }
}
