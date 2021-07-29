using OfficeOpenXml;
using OfficeOpenXml.Style;
using ServerHealthReport.Models;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Threading.Tasks;

namespace ServerHealthReport
{
    class Program
    {
        static async Task Main(string[] args)
        {

            //EPPlus Licensing
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var file = new FileInfo("___ServerList.xlsx");

            var serversList = await GetServerListAsync(file).ConfigureAwait(false);
            var allInfo = GetServerInfo(serversList);

            await GenerateExcelAsync(allInfo).ConfigureAwait(false);

        }

        private static async Task GenerateExcelAsync(IEnumerable<ServerInformation> serverInfo)
        {
            try
            {
                var outputFolder = "___HealthReports";

                Directory.CreateDirectory(outputFolder);
                //var file = new FileInfo($@"{outputFolder}\{DateTime.Today.Date.ToShortDateString().Replace("/", string.Empty)}.xlsx");
                var file = new FileInfo($@"{outputFolder}\{DateTime.Today:yyyy-MMMM-dd}.xlsx");

                await SaveExcelFileAsync(serverInfo, file).ConfigureAwait(false);


                //foreach (var serverInformation in serverInfo)
                //{


                //}
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        private static async Task SaveExcelFileAsync(IEnumerable<ServerInformation> serverInfo, FileInfo file)
        {
            DeleteIfExists(file);

            using var package = new ExcelPackage(file);

            var workSheet = package.Workbook.Worksheets.Add("HealthReport");




            //var range = workSheet.Cells["A1"].LoadFromCollection(serverInfo, true);

            //range.AutoFitColumns();

            //await package.SaveAsync().ConfigureAwait(false);

            var i = 2;

            workSheet.Cells[1, 1].LoadFromText("Server Name, " +
                                                       "Server IP," +
                                                       "CPU Status," +
                                                       "memory Status," +
                                                       "machine Status," +
                                                       "Disk Driver Letter," +
                                                       "Drive Name," +
                                                       "Total Size," +
                                                       "Used," +
                                                       "Free," +
                                                       "Free Percentage");




            foreach (var serverInformation in serverInfo)
            {


                var cpuStatus = $"CPU: with {serverInformation.NumberOfCores} Cores avg Load {serverInformation.ProcessorLoadAverage}%";
                var memoryStatus =
                    $"RAM: {serverInformation.TotalVisibleMemory:F} GB with {serverInformation.FreeMemoryPercentage:F}% Free";
                var machineStatus = $"UPTIME: {serverInformation.ServerUptime}";


                if (string.IsNullOrWhiteSpace(serverInformation.Error))
                {
                    foreach (var diskInformation in serverInformation.DiskInfo)
                    {



                        workSheet.Cells[i, 1].LoadFromText($"{serverInformation.ServerName}, " +
                                                           $"{serverInformation.ServerIp}," +
                                                           $"{cpuStatus}," +
                                                           $"{memoryStatus}," +
                                                           $"{machineStatus}," +
                                                           $"{diskInformation.DiskDriverLetter}," +
                                                           $"{diskInformation.DriveName}," +
                                                           $"{diskInformation.TotalSize} GB," +
                                                           $"{diskInformation.Used} GB," +
                                                           $"{diskInformation.Free} GB," +
                                                           $"{diskInformation.FreePercentage:F} %");

                        workSheet.Cells[i, 11].Style.Numberformat.Format = "#0.00%";

                        i++;
                    }
                }
                else
                {
                    workSheet.Cells[i, 2, i, 11].Merge = true;

                    //workSheet.Cells[i, 1].Value = serverInformation.ServerName;
                    //workSheet.Cells[i, 2].Value = serverInformation.Error;

                    workSheet.Cells[i, 1, i, 11].Style.Font.Color.SetColor(Color.Yellow);

                    workSheet.Cells[i, 1, i, 11].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    workSheet.Cells[i, 1, i, 11].Style.Fill.BackgroundColor.SetColor(Color.Red);


                    workSheet.Cells[i, 1].LoadFromText($"{serverInformation.ServerName}," +
                                                       $"{serverInformation.Error}");


                }



            }

            await package.SaveAsync().ConfigureAwait(false);




        }

        private static void DeleteIfExists(FileSystemInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        private static IEnumerable<ServerInformation> GetServerInfo(IEnumerable<string> serversList)
        {
            List<ServerInformation> serverInfoList = new();


            foreach (var server in serversList)
            {
                var serverInfo = GetServerInfo(server);
                serverInfoList.Add(serverInfo);
            }

            return serverInfoList;
        }

        /// <summary>
        /// Get Server List from the Excel File
        /// </summary>
        /// <param name="file">Server List Excel File</param>
        /// <returns></returns>
        private static async Task<List<string>> GetServerListAsync(FileInfo file)
        {
            List<string> serverList = new();

            using var package = new ExcelPackage(file);

            await package.LoadAsync(file).ConfigureAwait(false);
            var workSheet = package.Workbook.Worksheets[0];

            var row = 2;
            var column = 1;


            while (string.IsNullOrWhiteSpace(workSheet.Cells[row, column].Value?.ToString()) == false)
            {
                serverList.Add(workSheet.Cells[row, column].Value.ToString());
                row++;
            }

            return serverList;
        }

        private static ServerInformation GetServerInfo(string serverName)
        {
            Console.WriteLine($"Processing {serverName}");
            ServerInformation serverInfo = new();

            try
            {
                var scope = new ManagementScope($@"\\{serverName}\root\cimv2");
                scope.Connect();


                var osQuery = new ObjectQuery("SELECT * FROM Win32_OperatingSystem");
                var diskQuery = new SelectQuery("Select * from Win32_LogicalDisk  WHERE DriveType=3");
                var processorQuery = new SelectQuery("Select * from Win32_Processor");
                var networkQuery = new SelectQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = 'TRUE'");
                var computerSystemQuery = new SelectQuery("Select * from Win32_ComputerSystem");


                var osSearcher = new ManagementObjectSearcher(scope, osQuery);
                var osQueryCollection = osSearcher.Get();

                var diskSearcher = new ManagementObjectSearcher(scope, diskQuery);
                var diskQueryCollection = diskSearcher.Get();

                var processorSearcher = new ManagementObjectSearcher(scope, processorQuery);
                var processorQueryCollection = processorSearcher.Get();

                var networkSearcher = new ManagementObjectSearcher(scope, networkQuery);
                var networkQueryCollection = networkSearcher.Get();

                var computerSystemSearcher = new ManagementObjectSearcher(scope, computerSystemQuery);
                var computerSystemQueryCollection = computerSystemSearcher.Get();



                serverInfo.ServerName = serverName;


                //Get Server IP
                var serverIpAddress = networkQueryCollection
                    .Cast<ManagementObject>()
                    .SelectMany(o => (string[])(o["IPAddress"]))
                    .FirstOrDefault(a => a.Contains('.'));
                serverInfo.ServerIp = serverIpAddress;



                //Get Number of Processor
                foreach (var computerSystemQueryResult in computerSystemQueryCollection)
                {
                    serverInfo.NumberOfProcessors = computerSystemQueryResult["NumberOfProcessors"]?.ToString();
                }



                //Get Number of Cores and load Average
                var cores = 0;
                var loadPercentage = new List<int>();
                foreach (var processorQueryResult in processorQueryCollection)
                {
                    cores += int.Parse(processorQueryResult["NumberOfCores"].ToString());

                    //Load percentage
                    loadPercentage.Add(int.Parse(processorQueryResult["LoadPercentage"].ToString()));

                }
                serverInfo.NumberOfCores = cores.ToString();
                serverInfo.ProcessorLoadAverage = loadPercentage.AsQueryable().Average().ToString();




                foreach (ManagementObject osQueryResult in osQueryCollection)
                {
                    serverInfo.TotalVisibleMemory = decimal.Parse(osQueryResult["TotalVisibleMemorySize"].ToString()) / 1024 / 1024;
                    serverInfo.FreePhysicalMemory = decimal.Parse(osQueryResult["FreePhysicalMemory"].ToString()) / 1024 / 1024;
                    serverInfo.FreeMemoryPercentage =
                        (serverInfo.FreePhysicalMemory / serverInfo.TotalVisibleMemory) * 100;


                    serverInfo.LastBootTime = ManagementDateTimeConverter.ToDateTime(osQueryResult["LastBootUpTime"].ToString());

                    var timeGap = DateTime.Now - serverInfo.LastBootTime;

                    var days = timeGap.Days;
                    var hours = timeGap.Hours;
                    var minutes = timeGap.Minutes;
                    var seconds = timeGap.Seconds;


                    string day = string.Empty;
                    string hour = string.Empty;
                    string minute = string.Empty;
                    string second = string.Empty;


                    day = days == 1 ? "1 day " : $"{days} days ";
                    hour = hours == 1 ? "1 hour " : $"{hours} hours ";
                    minute = minutes == 1 ? "1 minute " : $"{minutes} minutes ";
                    second = seconds == 1 ? "1 second " : $"{seconds} seconds ";

                    serverInfo.ServerUptime = $"{day} {hour} {minute} {second}";
                }




                //DISK SECTION STARTS
                foreach (var disksResult in diskQueryCollection)
                {
                    var free = Math.Round((decimal.Parse(disksResult["FreeSpace"].ToString()) / 1073741824), 2);
                    var total = Math.Round((decimal.Parse(disksResult["Size"].ToString()) / 1073741824), 2);

                    serverInfo.DiskInfo.Add(new DiskInformation()
                    {
                        DiskDriverLetter = disksResult["DeviceID"].ToString(),
                        DriveName = disksResult["VolumeName"].ToString(),
                        Free = free,
                        TotalSize = total,
                        Used = total - free,
                        FreePercentage = (free / total) * 100
                    });
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error {serverName} - {ex.Message}");
                serverInfo.ServerName = serverName;
                serverInfo.Error = ex.Message;
            }

            Console.WriteLine($"Processed {serverName}");

            return serverInfo;


        }
    }
}
