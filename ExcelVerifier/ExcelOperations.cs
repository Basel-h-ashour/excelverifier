using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System.Collections.Generic;

namespace ExcelVerifier
{
    class ExcelOperations
    {
        public static OleDbConnection GetConnection(string filename, bool openIt)
        {
            // if your data has no header row, change HDR=NO
            var c = new OleDbConnection($"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='{filename}';Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\" ");
            if (openIt)
                c.Open();
            return c;
        }

        public static DataSet GetExcelFileAsDataSet(OleDbConnection conn)
        {
            var sheets = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new[] { default, default, default, "TABLE" });
            var ds = new DataSet();
            foreach (DataRow r in sheets.Rows)
            {
                if (r["TABLE_NAME"].ToString().EndsWith("_") || r["TABLE_NAME"].ToString().EndsWith("_filterdatabase") || !r["TABLE_NAME"].ToString().EndsWith("$")) continue;
                ds.Tables.Add(GetExcelSheetAsDataTable(conn, r["TABLE_NAME"].ToString()));
            }
            return ds;
        }

        public static DataTable GetExcelSheetAsDataTable(OleDbConnection conn, string sheetName)
        {
            using (var da = new OleDbDataAdapter($"select * from [{sheetName}]", conn))
            {
                var dt = new DataTable() { TableName = sheetName.TrimEnd('$') };
                da.Fill(dt);
                return dt;
            }
        }

        public static bool FileIsOpened(FileInfo file)
        {
            try
            {
                using (FileStream stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }

            //file is not locked
            return false;
        }
    }
}
