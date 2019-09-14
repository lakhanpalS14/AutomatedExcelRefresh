/*
  
Project Name  : RefreshPivots

Author       : Shivam Lakhanpal

Version      : 1.1
*/
using System;
using System.IO;

using System.Data;

using System.Linq;
using System.Net.NetworkInformation;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;  // reference to excel 12 object library!
using System.Net.Sockets;


namespace ExcelPivotRefresh
{
    class Program
    {
        static void Main(string[] args)
        {

            if (checkPing())
            {
                FileLocation fileLocation = null;
                try
                {
                    using (StreamReader sr = new StreamReader(Directory.GetCurrentDirectory() + @"\Input\Config.json"))
                    {
                        fileLocation = JsonConvert.DeserializeObject<FileLocation>(sr.ReadToEnd());
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                }

                Console.WriteLine("Connection Successful");
                string str = "";
                var Files = Directory.EnumerateFiles(fileLocation.SourceLocation, "*.*", SearchOption.AllDirectories)
                .Where(s => s.EndsWith(".xlsx"));
                Excel.Application xlApp = new Excel.Application();
                bool refreshStatus = true;
                foreach (string filePath in Files)
                {
                    DirectoryInfo file = new DirectoryInfo(filePath);
                    Excel.Workbook wb = xlApp.Workbooks.Open(file.FullName);
                    xlApp.DisplayAlerts = false;
                    xlApp.Visible = false;
                    Console.WriteLine("Refreshing : " + file.FullName);
                    Excel.Sheets excelSheets = wb.Worksheets;
                    foreach (Excel.Worksheet workSheet in excelSheets)
                    {
                        Console.WriteLine("SheetName: " + workSheet.Name);
                        Excel.PivotTables pivotTables = workSheet.PivotTables();
                        if (pivotTables.Count > 0)
                            foreach (Excel.PivotTable pivotTable in pivotTables)
                            {
                                Console.WriteLine(pivotTable.RefreshDate);
                                int attempts = 0;
                                Console.WriteLine("Refreshing " + pivotTable.Name);
                                for (; attempts < 3; attempts++)
                                {
                                    refreshStatus = pivotTable.RefreshTable();
                                    Console.WriteLine(refreshStatus);
                                    if (refreshStatus == true)
                                        break;
                                    else
                                        Console.WriteLine("Failed!! Reattempting...");
                                }
                                if (attempts == 3)
                                    Console.WriteLine("All attempts exhausted! Failed.");
                                else
                                    Console.WriteLine(pivotTable.RefreshDate);
                            }
                        else
                            Console.WriteLine("No Pivot Found in the Sheet!!");
                    }
                    if (refreshStatus)
                    {
                        Console.WriteLine("Refreshed :" + file.Name);
                        Console.WriteLine("Saving " + file.Name);
                        wb.SaveAs(fileLocation.DestinationLocation + file.Name);
                    }
                    wb.Close();
                    xlApp.Quit();
                    str = str + ", " + file.Name;
                }
                Console.WriteLine("Press Enter to Continue");
                Console.ReadLine();
            }
            else
                Console.WriteLine("Connection Unsuccessful");
        }
        public static bool checkPing()
        {

            var ping = new Ping();
            var reply = ping.Send("10.158.72.98", 5 * 1000); // 1 minute time out (in ms)
            if (reply.Status == IPStatus.Success)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        private void CheckEnter(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                // Enter key pressed
            }
        }

    }



}

