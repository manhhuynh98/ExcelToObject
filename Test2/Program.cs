using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
namespace Test2
{
    class Program
    {
        public static void addData(List<Information> items, int i)
        {
            string sql = "Server = localhost; Database = test; port = 3306; user id = root; password = ";
            MySqlConnection connection = new MySqlConnection(sql);
            connection.Open();
            MySqlCommand cmd = new MySqlCommand();
            cmd.CommandText = "INSERT INTO `infor`(`fullName`, `address1`, `address2`, `city`, `stateCode`, `zipCode`, `countryName`, `productName`, `productType`, `productColor`, `productSize`, `quantity`, `productID`) VALUES ('" + items[0].fullName + "','" + items[0].address1 + "','" + items[0].address2 + "','" + items[0].city + "','" + items[0].stateCode + "','" + items[0].zipCode + "','" + items[0].countryName + "','" + items[0].productName + "','" + items[0].productType + "','" + items[0].productColor + "','" + items[0].productSize + "','" + items[0].quantity + "','" + items[0].productID + "')";
            cmd.Connection = connection;

            //cmd.Parameters.AddWithValue("@fullName", items[0].fullName);
            //cmd.Parameters.AddWithValue("@address1", items[0].address1);
            //cmd.Parameters.AddWithValue("@address2", items[0].address2);
            //cmd.Parameters.AddWithValue("@city", items[0].city);
            //cmd.Parameters.AddWithValue("@stateCode", items[0].stateCode);
            //cmd.Parameters.AddWithValue("@zipCode", items[0].zipCode);
            //cmd.Parameters.AddWithValue("@countryName", items[0].countryName);
            //cmd.Parameters.AddWithValue("@productName", items[0].productName);
            //cmd.Parameters.AddWithValue("@productType", items[0].productType);
            //cmd.Parameters.AddWithValue("@productColor", items[0].productColor);
            //cmd.Parameters.AddWithValue("@productSize", items[0].productSize);
            //cmd.Parameters.AddWithValue("@quantity", items[0].quantity);
            //cmd.Parameters.AddWithValue("@productID", items[0].productID);

            cmd.ExecuteNonQuery();
            connection.Close();
            Console.WriteLine("Đã xong " + (i - 1));
            Console.WriteLine("Đã xong " + (i - 1));

        }
        public static void getExcelFile()
        {

            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"E:\Json\TEST-DATA.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            for (int i = 2; i <= rowCount; i++)
            {
                string fullName = "null";
                string address1 = "null";
                string address2 = "null";
                string city = "null";
                string stateCode = "null";
                string zipCode = "null";
                string countryName = "null";
                string productName = "null";
                string productType = "null";
                string productColor = "null";
                string productSize = "null";
                string quantity = "null";
                string productID = "null";

                if (xlRange.Cells[i, 1].Value2 != null) { fullName = xlRange.Cells[i, 1].Value2.ToString().Replace("'", @"\'"); }
                if (xlRange.Cells[i, 2].Value2 != null) { address1 = xlRange.Cells[i, 2].Value2.ToString().Replace("'", @"\'"); }
                if (xlRange.Cells[i, 3].Value2 != null) { address2 = xlRange.Cells[i, 3].Value2.ToString().Replace("'", @"\'"); }
                if (xlRange.Cells[i, 4].Value2 != null) { city = xlRange.Cells[i, 4].Value2.ToString().Replace("'", @"\'"); }
                if (xlRange.Cells[i, 5].Value2 != null) { stateCode = xlRange.Cells[i, 5].Value2.ToString().Replace("'", @"\'"); }
                if (xlRange.Cells[i, 6].Value2 != null) { zipCode = xlRange.Cells[i, 6].Value2.ToString().Replace("'", @"\'"); }
                if (xlRange.Cells[i, 7].Value2 != null) { countryName = xlRange.Cells[i, 7].Value2.ToString().Replace("'", @"\'"); }
                if (xlRange.Cells[i, 8].Value2 != null) { productName = xlRange.Cells[i, 8].Value2.ToString().Replace("'", @"\'"); }
                if (xlRange.Cells[i, 9].Value2 != null) { productType = xlRange.Cells[i, 9].Value2.ToString().Replace("'", @"\'"); }
                if (xlRange.Cells[i, 10].Value2 != null) { productColor = xlRange.Cells[i, 10].Value2.ToString().Replace("'", @"\'"); }
                if (xlRange.Cells[i, 11].Value2 != null) { productSize = xlRange.Cells[i, 11].Value2.ToString().Replace("'", @"\'"); }
                if (xlRange.Cells[i, 12].Value2 != null) { quantity = xlRange.Cells[i, 12].Value2.ToString().Replace("'", @"\'"); }
                if (xlRange.Cells[i, 13].Value2 != null) { productID = xlRange.Cells[i, 13].Value2.ToString().Replace("'", @"\'"); }

                List<Information> information = new List<Information>();
                information.Add(new Information(fullName, address1, address2, city, stateCode, zipCode, countryName, productName, productType, productColor, productSize, quantity, productID));
                addData(information, i);
               
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();

            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
        static void Main(string[] args)
        {
            getExcelFile();
            Console.ReadKey();
        }
    }
}
