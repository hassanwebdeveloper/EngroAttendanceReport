using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AttendanceReport
{
    public class CardHolderReportInfo
    {
        public string CardNumber { get; set; }

        public DateTime OccurrenceTime { get; set; }

        public string FirstName { get; set; }

        public string PNumber { get; set; }

        public string Crew { get; set; }

        public string Department { get; set; }

        public string Section { get; set; }

        public string Cadre { get; set; }

        public string CNICNumber { get; set; }

        public string Company { get; set; }

        public int NetNormalHours { get; set; }

        public int OverTimeHours { get; set; }

        public string CallOutFrom { get; set; }

        public string CallOutTo { get; set; }

        public int TotalCallOutHours { get; set; }

        public string ContractorName { get; set; }

        public string Category { get; set; }

        public DateTime BlockedTime { get; set; }

        public DateTime UnBlockedTime { get; set; }

        public string BlockedStatus { get; set; }
        
        public DateTime InTime { get; set; }

        public DateTime OutTime { get; set; }

        public DateTime CallOutInTime { get; set; }

        public DateTime CallOutOutTime { get; set; }

        public int FtItemId { get; set; }
    }
}
