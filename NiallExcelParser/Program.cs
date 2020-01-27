using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using ExcelDataReader;
using System.Threading.Tasks;
using System.Data;

namespace NiallExcelParser
{
    class Program
    {
        static void Main(string[] args)
        {
            using (FileStream fs = File.Open(args[0], FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(fs))
                {

                    //var dataset = reader.AsDataSet(reader);
                    var worksheet = reader.AsDataSet().Tables["MFWI-00074-SUD-F1"];
                    IList<System.Data.DataRow> rows = (from DataRow r in worksheet.Rows select r).ToList();

                    var rowsCount = rows.Count();
                    int columnCount = worksheet.Columns.Count;
                    var location = rows[1].ItemArray[1].ToString();


                    foreach (DataRow row in worksheet.Rows)
                    {
                        foreach(DataColumn column in worksheet.Columns)
                        {
                            var cellobj = row[column];
                            
                            var cell = row[column.ColumnName];
                        }

                    }
                }
            }
            
        }
    }
}
