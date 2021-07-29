using System;
using System.Collections.Generic;

namespace ServerHealthReport.Models
{
    public class ServerInformation
    {
        public string ServerName { get; set; }
        public string ServerIp { get; set; }

        public string NumberOfProcessors { get; set; }
        public string NumberOfCores { get; set; }
        public string ProcessorLoadAverage { get; set; }
        public decimal TotalVisibleMemory { get; set; }
        public decimal FreePhysicalMemory { get; set; }

        public decimal FreeMemoryPercentage { get; set; }

        public DateTime LastBootTime { get; set; }
        public string CpuStatus { get; set; }
        public string MemoryStatus { get; set; }
        public string MachineStatus { get; set; }
        public List<DiskInformation> DiskInfo { get; set; }

        public string Error { get; set; }
        public string ServerUptime { get; set; }


        public ServerInformation()
        {
            DiskInfo = new List<DiskInformation>();
        }
    }
}
