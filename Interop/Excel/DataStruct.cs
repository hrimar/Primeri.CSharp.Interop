using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel
{
    public class DataStruct
    {
        public List<DataRow> table = new List<DataRow>();

        public DataStruct()
        {
        }

        public void AddRow(string _fName, string _lName, string _age)
        {
            table.Add(new DataRow(_fName, _lName, _age));
        }

        public void PrintTable()
        {
            try
            {
                foreach (DataRow row in table)
                { 
                    Console.WriteLine(row.FirstName + " "+row.LastName+", "+row.Age);
                }
            }
            catch
            {
            }
        }
    }

    public class DataRow
    {
        private string _firstName = "";
        private string _lastName = "";
        private string _age = "";

        public DataRow(string firstName, string lastName, string age)
        {
            _firstName = firstName;
            _lastName = lastName;
            _age = age;

        }

        public string FirstName
        {
            set { _firstName = value;  }
            get { return _firstName;  }
        }

        public string LastName
        {
            set { _lastName = value; }
            get { return _lastName; }
        }

        public string Age
        {
            set { _age = value; }
            get { return _age; }
        }
    }
}
