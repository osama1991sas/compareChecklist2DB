using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace compareChecklist2DB.Classes
{
    internal class Helper
    {
        public static DataTable ReadExcelFile(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                DataSet result = reader.AsDataSet();
                return result.Tables[0]; // Get first sheet
            }
        }

        public static DataTable RemoveFirstThreeRows(DataTable table)
        {
            DataTable newTable = table.Clone(); // Clone structure

            for (int i = 3; i < table.Rows.Count; i++) // Skip first 3 rows
            {
                newTable.ImportRow(table.Rows[i]);
            }

            return newTable;
        }

        public static void SaveAsCsv(DataTable table, string filePath)
        {
            using (StreamWriter writer = new StreamWriter(filePath))
            {
                int rowCount = table.Rows.Count;
                for (int i = 1; i < rowCount; i++)
                {
                    writer.WriteLine(string.Join(",", table.Rows[i].ItemArray));
                }
            }
        }
    }
}
