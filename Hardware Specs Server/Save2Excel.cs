using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Hardware_Specs_GUI.Json;
using Newtonsoft.Json;
using OfficeOpenXml;

namespace Hardware_Specs_Server
{
    public class Response
    {
        public string CPUName { get; set; }
        public string Mob { get; set; }
        public string Systemname { get; set; }
        public string ThreadCount { get; set; }
        public string BaseClockSpeed { get; set; }
        public double MemoryCapacity { get; set; }
        public double MemorySpeed { get; set; }
        public string MemoryType { get; set; }
        public int MemoryDimms { get; set; }
        public string Cores { get; set; }
        public List<string> MACAddres { get; set; }
        public List<List<string>> NetworkAddresses { get; set; }
        public List<string> StorageNames { get; set; }
        public List<string> VideoName { get; set; }
        public List<double> Vram { get; set; }
    }

    public class Save2Excel
    {
        public static void Save(string path, string input)
        {
            Response response = input.FromJson<Response>();
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            ExcelWorksheet worksheet;
            using (ExcelPackage package = new ExcelPackage(path))
            {
                // Check if the sheet exists or not
                if (package.Workbook.Worksheets.Count == 0)
                {
                    worksheet = package.Workbook.Worksheets.Add("Computers");

                    // Excel column headers
                    worksheet.Cells[1, 1].Value = "System Name";
                    worksheet.Cells[1, 2].Value = "CPU Name";
                    worksheet.Cells[1, 3].Value = "Thread Count";
                    worksheet.Cells[1, 4].Value = "Base Clock Speed";
                    worksheet.Cells[1, 5].Value = "Memory Capacity";
                    worksheet.Cells[1, 6].Value = "Memory Speed";
                    worksheet.Cells[1, 7].Value = "Memory Type";
                    worksheet.Cells[1, 8].Value = "Memory Dimms";
                    worksheet.Cells[1, 9].Value = "Cores";
                    worksheet.Cells[1, 10].Value = "MAC Addresses";
                    worksheet.Cells[1, 11].Value = "Network Addresses";
                    worksheet.Cells[1, 12].Value = "Video Name";
                    worksheet.Cells[1, 13].Value = "VRAM";
                    worksheet.Cells[1, 14].Value = "Disk Model";
                    worksheet.Cells[1, 15].Value = "Motherboard";
                }
                else
                {
                    worksheet = package.Workbook.Worksheets.First();
                }

                // Load config values
                int rowIndex;
                for (rowIndex = 2; rowIndex < 200; rowIndex++)
                {
                    ExcelRange cell = worksheet.Cells[rowIndex, 1];
                    if (cell.Value == null || (string)cell.Value == response.Systemname)
                    {
                        break;
                    }
                }

                // Store Data
                worksheet.Cells[rowIndex, 1].Value = response.Systemname.ToString();
                worksheet.Cells[rowIndex, 2].Value = response.CPUName.ToString();
                worksheet.Cells[rowIndex, 3].Value = response.ThreadCount.ToString();
                worksheet.Cells[rowIndex, 4].Value = response.BaseClockSpeed.ToString();
                worksheet.Cells[rowIndex, 5].Value = response.MemoryCapacity.ToString();
                worksheet.Cells[rowIndex, 6].Value = response.MemorySpeed.ToString();
                worksheet.Cells[rowIndex, 7].Value = response.MemoryType.ToString();
                worksheet.Cells[rowIndex, 8].Value = response.MemoryDimms.ToString();
                worksheet.Cells[rowIndex, 9].Value = response.Cores.ToString();
                worksheet.Cells[rowIndex, 10].Value = string.Join(", ", response.MACAddres).ToString();
                string str = "";
                foreach (List<string> networks in response.NetworkAddresses)
                {
                    str += string.Join(", ", networks) + ", ";
                }
                worksheet.Cells[rowIndex, 11].Value = str;
                worksheet.Cells[rowIndex, 12].Value = string.Join(", ", response.VideoName);
                worksheet.Cells[rowIndex, 13].Value = string.Join(", ", response.Vram);
                worksheet.Cells[rowIndex, 14].Value = string.Join(", ", response.StorageNames);
                worksheet.Cells[rowIndex, 15].Value = response.Mob.ToString();
                
                // Save Excel file 
                package.SaveAs(new FileInfo(path));
            }

        }




    }
}
