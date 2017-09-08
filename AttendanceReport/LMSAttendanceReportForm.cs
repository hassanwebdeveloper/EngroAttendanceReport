using iText.Kernel.Colors;
using iText.Kernel.Events;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Layout;
using iText.Layout.Element;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AttendanceReport.EFERTDb;
using AttendanceReport.CCFTEvent;
using AttendanceReport.CCFTCentral;
using System.IO;

namespace AttendanceReport
{
    public partial class LMSAttendanceReportForm : Form
    {
        private Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>> mData = null;

        public LMSAttendanceReportForm()
        {
            InitializeComponent();

            EFERTDbUtility.UpdateDropDownFields(this.cbxDepartments, this.cbxSections, this.cbxCompany, this.cbxCadre);
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            this.mData = null;
            DateTime fromDate = this.dtpFromDate.Value.Date;
            DateTime toDate = this.dtpToDate.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);

            TimeSpan ndtStart = this.dtpNdtStart.Value.TimeOfDay;
            TimeSpan ndtEnd = this.dtpNdtEnd.Value.TimeOfDay;
            TimeSpan ndtLunchStart = this.dtpNdtLunchStart.Value.TimeOfDay;
            TimeSpan ndtLunchEnd = this.dtpNdtLunchEnd.Value.TimeOfDay;

            TimeSpan fdtStart = this.dtpFdtStart.Value.TimeOfDay;
            TimeSpan fdtEnd = this.dtpFdtEnd.Value.TimeOfDay;
            TimeSpan fdtLunchStart = this.dtpFdtLunchStart.Value.TimeOfDay;
            TimeSpan fdtLunchEnd = this.dtpFdtLunchEnd.Value.TimeOfDay;


            string filterByDepartment = this.cbxDepartments.Text.ToLower();
            string filterBySection = this.cbxSections.Text.ToLower();
            string filerByName = this.tbxName.Text.ToLower();
            string filterByCadre = this.cbxCadre.Text.ToLower();
            string filterByCompany = this.cbxCompany.Text.ToLower();
            string filterByCNIC = this.tbxCnic.Text;

            Dictionary<string, CardHolderReportInfo> cnicWiseReportInfo = new Dictionary<string, CardHolderReportInfo>();

            #region Actual Data

            List<CCFTEvent.Event> lstEvents = (from events in EFERTDbUtility.mCCFTEvent.Events
                                               where
                                                   events != null && (events.EventType == 20001 || events.EventType == 20003) &&
                                                   events.OccurrenceTime >= fromDate &&
                                                   events.OccurrenceTime < toDate
                                               select events).ToList();


            List<int> inIds = new List<int>();
            List<int> outIds = new List<int>();

            Dictionary<DateTime, Dictionary<int, List<CCFTEvent.Event>>> lstChlInEvents = new Dictionary<DateTime, Dictionary<int, List<CCFTEvent.Event>>>();
            Dictionary<DateTime, Dictionary<int, List<CCFTEvent.Event>>> lstChlOutEvents = new Dictionary<DateTime, Dictionary<int, List<CCFTEvent.Event>>>();

            Dictionary<int, Cardholder> inCardHolders = new Dictionary<int, Cardholder>();
            Dictionary<int, Cardholder> outCardHolders = new Dictionary<int, Cardholder>();

            Dictionary<int, List<CCFTEvent.Event>> dayWiseEvents = null;

            foreach (CCFTEvent.Event events in lstEvents)
            {
                if (events == null || events.RelatedItems == null)
                {
                    continue;
                }

                foreach (RelatedItem relatedItem in events.RelatedItems)
                {
                    if (relatedItem != null)
                    {
                        if (relatedItem.RelationCode == 0)
                        {
                            if (events.EventType == 20001)//In Events
                            {
                                inIds.Add(relatedItem.FTItemID);

                                if (lstChlInEvents.ContainsKey(events.OccurrenceTime.Date))
                                {
                                    if (lstChlInEvents[events.OccurrenceTime.Date].ContainsKey(relatedItem.FTItemID))
                                    {
                                        lstChlInEvents[events.OccurrenceTime.Date][relatedItem.FTItemID].Add(events);

                                    }
                                    else
                                    {

                                        lstChlInEvents[events.OccurrenceTime.Date].Add(relatedItem.FTItemID, new List<CCFTEvent.Event>() { events });
                                    }
                                }
                                else
                                {
                                    dayWiseEvents = new Dictionary<int, List<CCFTEvent.Event>>();
                                    dayWiseEvents.Add(relatedItem.FTItemID, new List<CCFTEvent.Event>() { events });

                                    lstChlInEvents.Add(events.OccurrenceTime.Date, dayWiseEvents);
                                }
                            }
                            else if(events.EventType == 20003)//Out events
                            {
                                outIds.Add(relatedItem.FTItemID);

                                if (lstChlOutEvents.ContainsKey(events.OccurrenceTime.Date))
                                {
                                    if (lstChlOutEvents[events.OccurrenceTime.Date].ContainsKey(relatedItem.FTItemID))
                                    {
                                        lstChlOutEvents[events.OccurrenceTime.Date][relatedItem.FTItemID].Add(events);
                                    }
                                    else
                                    {
                                        lstChlOutEvents[events.OccurrenceTime.Date].Add(relatedItem.FTItemID, new List<CCFTEvent.Event>() { events });
                                    }
                                }
                                else
                                {
                                    dayWiseEvents = new Dictionary<int, List<CCFTEvent.Event>>();
                                    dayWiseEvents.Add(relatedItem.FTItemID, new List<CCFTEvent.Event>() { events });

                                    lstChlOutEvents.Add(events.OccurrenceTime.Date, dayWiseEvents);
                                }
                            }
                            
                        }

                    }
                }
            }

            inCardHolders = (from chl in EFERTDbUtility.mCCFTCentral.Cardholders
                             where chl != null && inIds.Contains(chl.FTItemID)
                             select chl).Distinct().ToDictionary(ch => ch.FTItemID, ch => ch);

            List<string> strLstTempCards = (from chl in inCardHolders
                                            where chl.Value != null && (chl.Value.FirstName.ToLower().StartsWith("t-") || chl.Value.FirstName.ToLower().StartsWith("v-") || chl.Value.FirstName.ToLower().StartsWith("temporary-") || chl.Value.FirstName.ToLower().StartsWith("visitor-"))
                                            select chl.Value.LastName).ToList();

            List<CheckInAndOutInfo> filteredCheckIns = (from checkin in EFERTDbUtility.mEFERTDb.CheckedInInfos
                                                        where checkin != null && !checkin.CheckedIn && checkin.DateTimeIn >= fromDate && checkin.DateTimeIn < toDate &&
                                                            strLstTempCards.Contains(checkin.CardNumber) &&
                                                            (string.IsNullOrEmpty(filterByDepartment) ||
                                                                ((checkin.CardHolderInfos != null &&
                                                                checkin.CardHolderInfos.Department != null &&
                                                                checkin.CardHolderInfos.Department.DepartmentName.ToLower() == filterByDepartment) ||
                                                                (checkin.DailyCardHolders != null &&
                                                                checkin.DailyCardHolders.Department.ToLower() == filterByDepartment))) &&
                                                            (string.IsNullOrEmpty(filterBySection) ||
                                                                ((checkin.CardHolderInfos != null &&
                                                                checkin.CardHolderInfos.Section != null &&
                                                                checkin.CardHolderInfos.Section.SectionName.ToLower() == filterBySection) ||
                                                                (checkin.DailyCardHolders != null &&
                                                                checkin.DailyCardHolders.Section.ToLower() == filterBySection))) &&
                                                            (string.IsNullOrEmpty(filerByName) ||
                                                                ((checkin.CardHolderInfos != null &&
                                                                checkin.CardHolderInfos.FirstName.ToLower().Contains(filerByName)) ||
                                                                (checkin.DailyCardHolders != null &&
                                                                checkin.DailyCardHolders.FirstName.ToLower().Contains(filerByName)) ||
                                                                (checkin.Visitors != null &&
                                                                checkin.Visitors.FirstName.ToLower().Contains(filerByName)))) &&
                                                            (string.IsNullOrEmpty(filterByCadre) ||
                                                                ((checkin.CardHolderInfos != null &&
                                                                checkin.CardHolderInfos.Cadre != null &&
                                                                checkin.CardHolderInfos.Cadre.CadreName.ToLower() == filterByCadre) ||
                                                                (checkin.DailyCardHolders != null &&
                                                                checkin.DailyCardHolders.Cadre.ToLower() == filterByCadre))) &&
                                                            (string.IsNullOrEmpty(filterByCompany) ||
                                                                ((checkin.CardHolderInfos != null &&
                                                                checkin.CardHolderInfos.Company != null &&
                                                                checkin.CardHolderInfos.Company.CompanyName.ToLower() == filterByCompany) ||
                                                                (checkin.DailyCardHolders != null &&
                                                                checkin.DailyCardHolders.CompanyName.ToLower() == filterByCompany) ||
                                                                (checkin.Visitors != null &&
                                                                checkin.Visitors.CompanyName.ToLower() == filterByCompany))) &&
                                                            (string.IsNullOrEmpty(filterByCNIC) ||
                                                                ((checkin.CardHolderInfos != null &&
                                                                checkin.CardHolderInfos.CNICNumber == filterByCNIC) ||
                                                                (checkin.DailyCardHolders != null &&
                                                                checkin.DailyCardHolders.CNICNumber == filterByCNIC) ||
                                                                (checkin.Visitors != null &&
                                                                checkin.Visitors.CNICNumber == filterByCNIC)))
                                                        select checkin).ToList();


            outCardHolders = (from chl in EFERTDbUtility.mCCFTCentral.Cardholders
                              where chl != null && outIds.Contains(chl.FTItemID)
                              select chl).Distinct().ToDictionary(ch => ch.FTItemID, ch => ch);


            foreach (KeyValuePair<DateTime, Dictionary<int, List<CCFTEvent.Event>>> inEvent in lstChlInEvents)
            {
                DateTime date = inEvent.Key;
                if (inEvent.Value == null)
                {
                    continue;
                }

                foreach (KeyValuePair<int, List<CCFTEvent.Event>> chlWiseEvents in inEvent.Value)
                {
                    if (chlWiseEvents.Value == null || chlWiseEvents.Value.Count == 0 || !inCardHolders.ContainsKey(chlWiseEvents.Key))
                    {
                        continue;
                    }

                    int ftItemId = chlWiseEvents.Key;

                    Cardholder chl = inCardHolders[ftItemId];

                    if (chl == null)
                    {
                        continue;
                    }

                    bool isTempCard = chl.FirstName.ToLower().StartsWith("t-") || chl.FirstName.ToLower().StartsWith("v-") || chl.Value.FirstName.ToLower().StartsWith("temporary-") || chl.Value.FirstName.ToLower().StartsWith("visitor-");

                    List<CCFTEvent.Event> inEvents = chlWiseEvents.Value;

                    inEvents = inEvents.OrderBy(ev => ev.OccurrenceTime).ToList();

                    if (isTempCard)
                    {
                        string tempCardNumber = chl.LastName;

                        List<CheckInAndOutInfo> dateWiseCheckins = (from checkIn in filteredCheckIns
                                                                    where checkIn != null && checkIn.DateTimeIn.Date == date && checkIn.CardNumber == tempCardNumber
                                                                    select checkIn).ToList();

                        Dictionary<string, DateTime> dictInTime = new Dictionary<string, DateTime>();
                        Dictionary<string, DateTime> dictOutTime = new Dictionary<string, DateTime>();//dateWiseCheckIn.DateTimeOut;

                        Dictionary<string, DateTime> dictCallOutInTime = new Dictionary<string, DateTime>();
                        Dictionary<string, DateTime> dictCallOutOutTime = new Dictionary<string, DateTime>();

                        foreach (CheckInAndOutInfo dateWiseCheckIn in dateWiseCheckins)
                        {
                            string cnicNumber = dateWiseCheckIn.CNICNumber;
                            string firstName = dateWiseCheckIn.CardHolderInfos == null ? (dateWiseCheckIn.DailyCardHolders == null ? (dateWiseCheckIn.Visitors == null ? string.Empty : dateWiseCheckIn.Visitors.FirstName) : dateWiseCheckIn.DailyCardHolders.FirstName) : dateWiseCheckIn.CardHolderInfos.FirstName;
                            string department = dateWiseCheckIn.CardHolderInfos == null ? (dateWiseCheckIn.DailyCardHolders == null ? string.Empty : dateWiseCheckIn.DailyCardHolders.Department) : dateWiseCheckIn.CardHolderInfos.Department?.DepartmentName;
                            string section = dateWiseCheckIn.CardHolderInfos == null ? (dateWiseCheckIn.DailyCardHolders == null ? string.Empty : dateWiseCheckIn.DailyCardHolders.Section) : dateWiseCheckIn.CardHolderInfos.Section?.SectionName;
                            string cadre = dateWiseCheckIn.CardHolderInfos == null ? (dateWiseCheckIn.DailyCardHolders == null ? string.Empty : dateWiseCheckIn.DailyCardHolders.Cadre) : dateWiseCheckIn.CardHolderInfos.Cadre?.CadreName;


                            DateTime inTime = dictInTime.ContainsKey(cnicNumber) ? dictInTime[cnicNumber] : DateTime.MaxValue;
                            DateTime outTime = dictOutTime.ContainsKey(cnicNumber) ? dictOutTime[cnicNumber] : DateTime.MaxValue;

                            DateTime callOutInTime = dictCallOutInTime.ContainsKey(cnicNumber) ? dictCallOutInTime[cnicNumber] : DateTime.MaxValue;
                            DateTime callOutOutTime = dictCallOutOutTime.ContainsKey(cnicNumber) ? dictCallOutOutTime[cnicNumber] : DateTime.MaxValue;

                            if (date.DayOfWeek == DayOfWeek.Friday)
                            {
                                if (dateWiseCheckIn.DateTimeIn.TimeOfDay < fdtEnd)
                                {
                                    if (inTime == DateTime.MaxValue)
                                    {
                                        inTime = dateWiseCheckIn.DateTimeIn;
                                        dictInTime.Add(cnicNumber, inTime);
                                    }
                                    else
                                    {
                                        if (dateWiseCheckIn.DateTimeIn.TimeOfDay < inTime.TimeOfDay)
                                        {
                                            inTime = dateWiseCheckIn.DateTimeIn;
                                            dictInTime[cnicNumber] = inTime;

                                        }
                                    }

                                }
                                else
                                {
                                    if (callOutInTime == DateTime.MaxValue)
                                    {
                                        callOutInTime = dateWiseCheckIn.DateTimeIn;
                                        dictCallOutInTime.Add(cnicNumber, inTime);
                                    }
                                    else
                                    {
                                        if (dateWiseCheckIn.DateTimeIn.TimeOfDay < callOutInTime.TimeOfDay)
                                        {
                                            callOutInTime = dateWiseCheckIn.DateTimeIn;
                                            dictCallOutInTime[cnicNumber] = inTime;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (dateWiseCheckIn.DateTimeIn.TimeOfDay < ndtEnd)
                                {
                                    if (inTime == DateTime.MaxValue)
                                    {
                                        inTime = dateWiseCheckIn.DateTimeIn;
                                        dictInTime.Add(cnicNumber, inTime);
                                    }
                                    else
                                    {
                                        if (dateWiseCheckIn.DateTimeIn.TimeOfDay < inTime.TimeOfDay)
                                        {
                                            inTime = dateWiseCheckIn.DateTimeIn;
                                            dictInTime[cnicNumber] = inTime;

                                        }
                                    }

                                }
                                else
                                {
                                    if (callOutInTime == DateTime.MaxValue)
                                    {
                                        callOutInTime = dateWiseCheckIn.DateTimeIn;
                                        dictCallOutInTime.Add(cnicNumber, inTime);
                                    }
                                    else
                                    {
                                        if (dateWiseCheckIn.DateTimeIn.TimeOfDay < callOutInTime.TimeOfDay)
                                        {
                                            callOutInTime = dateWiseCheckIn.DateTimeIn;
                                            dictCallOutInTime[cnicNumber] = inTime;
                                        }
                                    }
                                }
                            }

                            if (inTime == DateTime.MaxValue && callOutInTime == DateTime.MaxValue)
                            {
                                continue;
                            }

                            if (callOutInTime == DateTime.MaxValue)
                            {
                                if (dateWiseCheckIn.DateTimeOut.TimeOfDay > inTime.TimeOfDay)
                                {
                                    outTime = dateWiseCheckIn.DateTimeOut;

                                    if (dictOutTime.ContainsKey(cnicNumber))
                                    {
                                        dictOutTime[cnicNumber] = outTime;
                                    }
                                    else
                                    {
                                        dictOutTime.Add(cnicNumber, outTime);
                                    }
                                }
                            }
                            else
                            {
                                if (dateWiseCheckIn.DateTimeOut.TimeOfDay < callOutInTime.TimeOfDay)
                                {
                                    outTime = dateWiseCheckIn.DateTimeOut;

                                    if (dictOutTime.ContainsKey(cnicNumber))
                                    {
                                        dictOutTime[cnicNumber] = outTime;
                                    }
                                    else
                                    {
                                        dictOutTime.Add(cnicNumber, outTime);
                                    }
                                }
                                else
                                {
                                    callOutOutTime = dateWiseCheckIn.DateTimeOut;

                                    if (dictCallOutOutTime.ContainsKey(cnicNumber))
                                    {
                                        dictCallOutOutTime[cnicNumber] = callOutOutTime;
                                    }
                                    else
                                    {
                                        dictCallOutOutTime.Add(cnicNumber, callOutOutTime);
                                    }
                                }
                            }

                            if (outTime == DateTime.MaxValue && callOutOutTime == DateTime.MaxValue)
                            {
                                continue;
                            }

                            if (cnicWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                            {
                                DateTime prevInTime = cnicWiseReportInfo[cnicNumber].InTime;
                                DateTime prevOutTime = cnicWiseReportInfo[cnicNumber].OutTime;

                                DateTime prevCallOutInTime = cnicWiseReportInfo[cnicNumber].CallOutInTime;
                                DateTime prevCallOutOutTime = cnicWiseReportInfo[cnicNumber].CallOutOutTime;

                                if (date.DayOfWeek == DayOfWeek.Friday)
                                {
                                    if (inTime.TimeOfDay < fdtEnd)
                                    {
                                        if (inTime.TimeOfDay > prevOutTime.TimeOfDay)
                                        {
                                            inTime = prevInTime;
                                        }

                                        if (outTime.TimeOfDay < prevOutTime.TimeOfDay)
                                        {
                                            outTime = prevOutTime;
                                        }
                                    }
                                    else
                                    {
                                        if (inTime.TimeOfDay > prevCallOutOutTime.TimeOfDay)
                                        {
                                            callOutInTime = prevCallOutInTime;
                                        }

                                        if (callOutOutTime.TimeOfDay < prevCallOutOutTime.TimeOfDay)
                                        {
                                            callOutOutTime = prevCallOutOutTime;
                                        }
                                    }
                                }
                                else
                                {
                                    if (inTime.TimeOfDay < ndtEnd)
                                    {
                                        if (inTime.TimeOfDay > prevOutTime.TimeOfDay)
                                        {
                                            inTime = prevInTime;
                                        }

                                        if (outTime.TimeOfDay < prevOutTime.TimeOfDay)
                                        {
                                            outTime = prevOutTime;
                                        }
                                    }
                                    else
                                    {
                                        if (inTime.TimeOfDay > prevCallOutOutTime.TimeOfDay)
                                        {
                                            callOutInTime = prevCallOutInTime;
                                        }

                                        if (callOutOutTime.TimeOfDay < prevCallOutOutTime.TimeOfDay)
                                        {
                                            callOutOutTime = prevCallOutOutTime;
                                        }
                                    }
                                }

                            }

                            int netNormalHours = 0;
                            int otHours = 0;
                            int callOutHours = 0;
                            string callOutFromHours = string.Empty;
                            string callOutToHours = string.Empty;
                            int lunchHours = 0;

                            if (date.DayOfWeek == DayOfWeek.Friday)
                            {
                                lunchHours = (fdtLunchStart - fdtLunchEnd).Hours;

                                if (inTime.TimeOfDay < fdtLunchStart)
                                {
                                    if (outTime.TimeOfDay < fdtLunchEnd)
                                    {
                                        netNormalHours = (fdtLunchStart - inTime.TimeOfDay).Hours;
                                    }
                                    else
                                    {
                                        if (outTime.TimeOfDay <= fdtEnd)
                                        {
                                            netNormalHours = (outTime.TimeOfDay - inTime.TimeOfDay).Hours - lunchHours;
                                        }
                                        else
                                        {
                                            netNormalHours = (fdtEnd - inTime.TimeOfDay).Hours - lunchHours;
                                            otHours = (outTime.TimeOfDay - fdtEnd).Hours;
                                        }

                                    }

                                }
                                else
                                {
                                    if (inTime.TimeOfDay < fdtLunchEnd)
                                    {
                                        if (outTime.TimeOfDay > fdtLunchEnd)
                                        {
                                            if (outTime.TimeOfDay <= fdtEnd)
                                            {
                                                netNormalHours = (outTime.TimeOfDay - fdtLunchEnd).Hours;
                                            }
                                            else
                                            {
                                                netNormalHours = (fdtEnd - fdtLunchEnd).Hours - lunchHours;
                                                otHours = (outTime.TimeOfDay - fdtEnd).Hours;
                                            }
                                        }

                                    }
                                    else
                                    {
                                        if (outTime.TimeOfDay <= fdtEnd)
                                        {
                                            netNormalHours = (outTime.TimeOfDay - inTime.TimeOfDay).Hours;
                                        }
                                        else
                                        {
                                            netNormalHours = (fdtEnd - inTime.TimeOfDay).Hours - lunchHours;
                                            otHours = (outTime.TimeOfDay - fdtEnd).Hours;
                                        }

                                    }
                                }
                            }
                            else
                            {
                                lunchHours = (ndtLunchStart - ndtLunchEnd).Hours;

                                if (inTime.TimeOfDay < ndtLunchStart)
                                {
                                    if (outTime.TimeOfDay < ndtLunchEnd)
                                    {
                                        netNormalHours = (ndtLunchStart - inTime.TimeOfDay).Hours;
                                    }
                                    else
                                    {
                                        if (outTime.TimeOfDay <= ndtEnd)
                                        {
                                            netNormalHours = (outTime.TimeOfDay - inTime.TimeOfDay).Hours - lunchHours;
                                        }
                                        else
                                        {
                                            netNormalHours = (ndtEnd - inTime.TimeOfDay).Hours - lunchHours;
                                            otHours = (outTime.TimeOfDay - ndtEnd).Hours;
                                        }

                                    }

                                }
                                else
                                {
                                    if (inTime.TimeOfDay < ndtLunchEnd)
                                    {
                                        if (outTime.TimeOfDay > ndtLunchEnd)
                                        {
                                            if (outTime.TimeOfDay <= ndtEnd)
                                            {
                                                netNormalHours = (outTime.TimeOfDay - ndtLunchEnd).Hours;
                                            }
                                            else
                                            {
                                                netNormalHours = (ndtEnd - ndtLunchEnd).Hours - lunchHours;
                                                otHours = (outTime.TimeOfDay - ndtEnd).Hours;
                                            }
                                        }

                                    }
                                    else
                                    {
                                        if (outTime.TimeOfDay <= ndtEnd)
                                        {
                                            netNormalHours = (outTime.TimeOfDay - inTime.TimeOfDay).Hours;
                                        }
                                        else
                                        {
                                            netNormalHours = (ndtEnd - inTime.TimeOfDay).Hours - lunchHours;
                                            otHours = (outTime.TimeOfDay - ndtEnd).Hours;
                                        }

                                    }
                                }
                            }

                            if (callOutInTime != DateTime.MaxValue && callOutOutTime != DateTime.MaxValue)
                            {
                                callOutHours = (callOutOutTime.TimeOfDay - callOutInTime.TimeOfDay).Hours;
                                callOutFromHours = callOutInTime.ToString("HH:mm");
                                callOutToHours = callOutOutTime.ToString("HH:mm");
                            }

                            if (cnicWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                            {
                                CardHolderReportInfo reportInfo = cnicWiseReportInfo[cnicNumber + "^" + date.ToString()];

                                if (reportInfo != null)
                                {
                                    if (reportInfo.NetNormalHours < netNormalHours)
                                    {
                                        reportInfo.InTime = inTime;
                                        reportInfo.OutTime = outTime;

                                        reportInfo.NetNormalHours = netNormalHours;
                                    }

                                    if (reportInfo.OverTimeHours < otHours)
                                    {
                                        reportInfo.InTime = inTime;
                                        reportInfo.OutTime = outTime;

                                        reportInfo.OverTimeHours = otHours;
                                    }

                                    if (reportInfo.TotalCallOutHours < callOutHours)
                                    {
                                        reportInfo.CallOutInTime = callOutInTime;
                                        reportInfo.CallOutOutTime = callOutOutTime;

                                        reportInfo.TotalCallOutHours = callOutHours;
                                        reportInfo.CallOutFrom = callOutFromHours;
                                        reportInfo.CallOutTo = callOutToHours;
                                    }
                                }
                            }
                            else
                            {
                                cnicWiseReportInfo.Add(cnicNumber + "^" + date.ToString(), new CardHolderReportInfo()
                                {
                                    OccurrenceTime = date,
                                    FirstName = firstName,
                                    CNICNumber = cnicNumber,
                                    Department = department,
                                    Section = section,
                                    Cadre = cadre,
                                    NetNormalHours = netNormalHours,
                                    OverTimeHours = otHours,
                                    CallOutFrom = callOutFromHours,
                                    CallOutTo = callOutToHours,
                                    TotalCallOutHours = callOutHours,
                                    InTime = inTime,
                                    OutTime = outTime,
                                    CallOutInTime = callOutInTime,
                                    CallOutOutTime = callOutOutTime
                                });
                            }
                        }
                    }
                    else
                    {
                        if (!lstChlOutEvents.ContainsKey(date) ||
                            lstChlOutEvents[date] == null ||
                            !lstChlOutEvents[date].ContainsKey(ftItemId) ||
                            lstChlOutEvents[date][ftItemId] == null ||
                            lstChlOutEvents[date][ftItemId].Count == 0)
                        {
                            continue;
                        }

                        List<CCFTEvent.Event> outEvents = lstChlOutEvents[date][ftItemId];

                        outEvents = outEvents.OrderBy(ev => ev.OccurrenceTime).ToList();

                        int pNumber = chl.PersonalDataIntegers == null || chl.PersonalDataIntegers.Count == 0 ? 0 : Convert.ToInt32(chl.PersonalDataIntegers.ElementAt(0).Value);
                        string cnicNumber = chl.PersonalDataStrings?.ToList()?.Find(pds => pds.PersonalDataFieldID == 5051)?.Value;
                        string department = chl.PersonalDataStrings?.ToList()?.Find(pds => pds.PersonalDataFieldID == 5043)?.Value;
                        string section = chl.PersonalDataStrings?.ToList()?.Find(pds => pds.PersonalDataFieldID == 12951)?.Value;
                        string cadre = chl.PersonalDataStrings?.ToList()?.Find(pds => pds.PersonalDataFieldID == 12952)?.Value;
                        string company = chl.PersonalDataStrings?.ToList()?.Find(pds => pds.PersonalDataFieldID == 5059)?.Value;

                        //Filter By Department
                        if (string.IsNullOrEmpty(department) || !string.IsNullOrEmpty(filterByDepartment) && department.ToLower() != filterByDepartment.ToLower())
                        {
                            continue;
                        }

                        //Filter By Section
                        if (string.IsNullOrEmpty(section) || !string.IsNullOrEmpty(filterBySection) && section.ToLower() != filterBySection.ToLower())
                        {
                            continue;
                        }

                        //Filter By Cadre
                        if (string.IsNullOrEmpty(cadre) || !string.IsNullOrEmpty(filterByCadre) && section.ToLower() != filterByCadre.ToLower())
                        {
                            continue;
                        }

                        //Filter By Company
                        if (!string.IsNullOrEmpty(filterByCompany) && company.ToLower() != filterByCompany.ToLower())
                        {
                            continue;
                        }

                        //Filter By CNIC
                        if (string.IsNullOrEmpty(cnicNumber) || !string.IsNullOrEmpty(filterByCNIC) && cnicNumber != filterByCNIC)
                        {
                            continue;
                        }

                        //Filter By Name
                        if (!string.IsNullOrEmpty(filerByName) && !chl.FirstName.Contains(filerByName))
                        {
                            continue;
                        }

                        DateTime inTime = DateTime.MaxValue;
                        DateTime outTime = DateTime.MaxValue;

                        DateTime callOutInTime = DateTime.MaxValue;
                        DateTime callOutOutTime = DateTime.MaxValue;

                        foreach (CCFTEvent.Event ev in inEvents)
                        {
                            DateTime inDateTime = ev.OccurrenceTime.AddHours(5);

                            if (date.DayOfWeek == DayOfWeek.Friday)
                            {
                                if (inDateTime.TimeOfDay < fdtEnd)
                                {
                                    if (inTime == DateTime.MaxValue)
                                    {
                                        inTime = inDateTime;
                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < inTime.TimeOfDay)
                                        {
                                            inTime = inDateTime;
                                        }
                                    }

                                }
                                else
                                {
                                    if (callOutInTime == DateTime.MaxValue)
                                    {
                                        callOutInTime = inDateTime;
                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < callOutInTime.TimeOfDay)
                                        {
                                            callOutInTime = inDateTime;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (inDateTime.TimeOfDay < ndtEnd)
                                {
                                    if (inTime == DateTime.MaxValue)
                                    {
                                        inTime = inDateTime;
                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < inTime.TimeOfDay)
                                        {
                                            inTime = inDateTime;
                                        }
                                    }

                                }
                                else
                                {
                                    if (callOutInTime == DateTime.MaxValue)
                                    {
                                        callOutInTime = inDateTime;
                                    }
                                    else
                                    {
                                        if (inDateTime.TimeOfDay < callOutInTime.TimeOfDay)
                                        {
                                            callOutInTime = inDateTime;
                                        }
                                    }
                                }
                            }
                        }

                        if (inTime == DateTime.MaxValue && callOutInTime == DateTime.MaxValue)
                        {
                            continue;
                        }

                        foreach (CCFTEvent.Event ev in outEvents)
                        {
                            DateTime outDateTime = ev.OccurrenceTime.AddHours(5);

                            if (callOutInTime == DateTime.MaxValue)
                            {
                                if (outDateTime.TimeOfDay > inTime.TimeOfDay)
                                {
                                    outTime = outDateTime;
                                }

                            }
                            else
                            {
                                if (outDateTime.TimeOfDay < callOutInTime.TimeOfDay)
                                {
                                    outTime = outDateTime;
                                }
                                else
                                {
                                    callOutOutTime = outDateTime;
                                }
                            }

                        }

                        if (outTime == DateTime.MaxValue && callOutOutTime == DateTime.MaxValue)
                        {
                            continue;
                        }

                        if (cnicWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                        {
                            DateTime prevInTime = cnicWiseReportInfo[cnicNumber].InTime;
                            DateTime prevOutTime = cnicWiseReportInfo[cnicNumber].OutTime;

                            DateTime prevCallOutInTime = cnicWiseReportInfo[cnicNumber].CallOutInTime;
                            DateTime prevCallOutOutTime = cnicWiseReportInfo[cnicNumber].CallOutOutTime;

                            if (date.DayOfWeek == DayOfWeek.Friday)
                            {
                                if (inTime.TimeOfDay < fdtEnd)
                                {
                                    if (inTime.TimeOfDay > prevOutTime.TimeOfDay)
                                    {
                                        inTime = prevInTime;
                                    }

                                    if (outTime.TimeOfDay < prevOutTime.TimeOfDay)
                                    {
                                        outTime = prevOutTime;
                                    }
                                }
                                else
                                {
                                    if (inTime.TimeOfDay > prevCallOutOutTime.TimeOfDay)
                                    {
                                        callOutInTime = prevCallOutInTime;
                                    }

                                    if (callOutOutTime.TimeOfDay < prevCallOutOutTime.TimeOfDay)
                                    {
                                        callOutOutTime = prevCallOutOutTime;
                                    }
                                }
                            }
                            else
                            {
                                if (inTime.TimeOfDay < ndtEnd)
                                {
                                    if (inTime.TimeOfDay > prevOutTime.TimeOfDay)
                                    {
                                        inTime = prevInTime;
                                    }

                                    if (outTime.TimeOfDay < prevOutTime.TimeOfDay)
                                    {
                                        outTime = prevOutTime;
                                    }
                                }
                                else
                                {
                                    if (inTime.TimeOfDay > prevCallOutOutTime.TimeOfDay)
                                    {
                                        callOutInTime = prevCallOutInTime;
                                    }

                                    if (callOutOutTime.TimeOfDay < prevCallOutOutTime.TimeOfDay)
                                    {
                                        callOutOutTime = prevCallOutOutTime;
                                    }
                                }
                            }

                        }

                        int netNormalHours = 0;
                        int otHours = 0;
                        int callOutHours = 0;
                        string callOutFromHours = string.Empty;
                        string callOutToHours = string.Empty;
                        int lunchHours = 0;

                        if (date.DayOfWeek == DayOfWeek.Friday)
                        {
                            lunchHours = (fdtLunchStart - fdtLunchEnd).Hours;

                            if (inTime.TimeOfDay < fdtLunchStart)
                            {
                                if (outTime.TimeOfDay < fdtLunchEnd)
                                {
                                    netNormalHours = (fdtLunchStart - inTime.TimeOfDay).Hours;
                                }
                                else
                                {
                                    if (outTime.TimeOfDay <= fdtEnd)
                                    {
                                        netNormalHours = (outTime.TimeOfDay - inTime.TimeOfDay).Hours - lunchHours;
                                    }
                                    else
                                    {
                                        netNormalHours = (fdtEnd - inTime.TimeOfDay).Hours - lunchHours;
                                        otHours = (outTime.TimeOfDay - fdtEnd).Hours;
                                    }

                                }

                            }
                            else
                            {
                                if (inTime.TimeOfDay < fdtLunchEnd)
                                {
                                    if (outTime.TimeOfDay > fdtLunchEnd)
                                    {
                                        if (outTime.TimeOfDay <= fdtEnd)
                                        {
                                            netNormalHours = (outTime.TimeOfDay - fdtLunchEnd).Hours;
                                        }
                                        else
                                        {
                                            netNormalHours = (fdtEnd - fdtLunchEnd).Hours - lunchHours;
                                            otHours = (outTime.TimeOfDay - fdtEnd).Hours;
                                        }
                                    }

                                }
                                else
                                {
                                    if (outTime.TimeOfDay <= fdtEnd)
                                    {
                                        netNormalHours = (outTime.TimeOfDay - inTime.TimeOfDay).Hours;
                                    }
                                    else
                                    {
                                        netNormalHours = (fdtEnd - inTime.TimeOfDay).Hours - lunchHours;
                                        otHours = (outTime.TimeOfDay - fdtEnd).Hours;
                                    }

                                }
                            }
                        }
                        else
                        {
                            lunchHours = (ndtLunchStart - ndtLunchEnd).Hours;

                            if (inTime.TimeOfDay < ndtLunchStart)
                            {
                                if (outTime.TimeOfDay < ndtLunchEnd)
                                {
                                    netNormalHours = (ndtLunchStart - inTime.TimeOfDay).Hours;
                                }
                                else
                                {
                                    if (outTime.TimeOfDay <= ndtEnd)
                                    {
                                        netNormalHours = (outTime.TimeOfDay - inTime.TimeOfDay).Hours - lunchHours;
                                    }
                                    else
                                    {
                                        netNormalHours = (ndtEnd - inTime.TimeOfDay).Hours - lunchHours;
                                        otHours = (outTime.TimeOfDay - ndtEnd).Hours;
                                    }

                                }

                            }
                            else
                            {
                                if (inTime.TimeOfDay < ndtLunchEnd)
                                {
                                    if (outTime.TimeOfDay > ndtLunchEnd)
                                    {
                                        if (outTime.TimeOfDay <= ndtEnd)
                                        {
                                            netNormalHours = (outTime.TimeOfDay - ndtLunchEnd).Hours;
                                        }
                                        else
                                        {
                                            netNormalHours = (ndtEnd - ndtLunchEnd).Hours - lunchHours;
                                            otHours = (outTime.TimeOfDay - ndtEnd).Hours;
                                        }
                                    }

                                }
                                else
                                {
                                    if (outTime.TimeOfDay <= ndtEnd)
                                    {
                                        netNormalHours = (outTime.TimeOfDay - inTime.TimeOfDay).Hours;
                                    }
                                    else
                                    {
                                        netNormalHours = (ndtEnd - inTime.TimeOfDay).Hours - lunchHours;
                                        otHours = (outTime.TimeOfDay - ndtEnd).Hours;
                                    }

                                }
                            }
                        }

                        if (callOutInTime != null && callOutOutTime != null)
                        {
                            callOutHours = (callOutOutTime.TimeOfDay - callOutInTime.TimeOfDay).Hours;
                            callOutFromHours = callOutInTime.ToString("HH:mm");
                            callOutToHours = callOutOutTime.ToString("HH:mm");
                        }

                        if (cnicWiseReportInfo.ContainsKey(cnicNumber + "^" + date.ToString()))
                        {
                            CardHolderReportInfo reportInfo = cnicWiseReportInfo[cnicNumber + "^" + date.ToString()];

                            if (reportInfo != null)
                            {
                                if (reportInfo.NetNormalHours < netNormalHours)
                                {
                                    reportInfo.InTime = inTime;
                                    reportInfo.OutTime = outTime;

                                    reportInfo.NetNormalHours = netNormalHours;
                                }

                                if (reportInfo.OverTimeHours < otHours)
                                {
                                    reportInfo.InTime = inTime;
                                    reportInfo.OutTime = outTime;

                                    reportInfo.OverTimeHours = otHours;
                                }

                                if (reportInfo.TotalCallOutHours < callOutHours)
                                {
                                    reportInfo.CallOutInTime = callOutInTime;
                                    reportInfo.CallOutOutTime = callOutOutTime;

                                    reportInfo.TotalCallOutHours = callOutHours;
                                    reportInfo.CallOutFrom = callOutFromHours;
                                    reportInfo.CallOutTo = callOutToHours;
                                }
                            }
                        }
                        else
                        {
                            cnicWiseReportInfo.Add(cnicNumber + "^" + date.ToString(), new CardHolderReportInfo()
                            {
                                OccurrenceTime = date,
                                FirstName = chl.FirstName,
                                CNICNumber = cnicNumber,
                                Department = department,
                                Section = section,
                                Cadre = cadre,
                                NetNormalHours = netNormalHours,
                                OverTimeHours = otHours,
                                CallOutFrom = callOutFromHours,
                                CallOutTo = callOutToHours,
                                TotalCallOutHours = callOutHours,
                                InTime = inTime,
                                OutTime = outTime,
                                CallOutInTime = callOutInTime,
                                CallOutOutTime = callOutOutTime
                            });
                        }

                    }
                }
            }

            #endregion

            #region Dummy Data

            //cnicWiseReportInfo.Add("11111-1111111-1",new CardHolderReportInfo() {
            //    OccurrenceTime = DateTime.Now.Date,
            //    FirstName = "Card Holder 1",
            //    CNICNumber = "11111-1111111-1",
            //    Department = "Department 1",
            //    Section = "Section 1",
            //    Cadre = "NMPT 1",
            //    NetNormalHours = 8,
            //    OverTimeHours = 2,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2
            //});

            //cnicWiseReportInfo.Add("11111-1111111-2", new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date,
            //    FirstName = "Card Holder 2",
            //    CNICNumber = "11111-1111111-2",
            //    Department = "Department 1",
            //    Section = "Section 1",
            //    Cadre = "NMPT 1",
            //    NetNormalHours = 6,
            //    OverTimeHours = 0,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2
            //});

            //cnicWiseReportInfo.Add("11111-1111111-3", new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date,
            //    FirstName = "Card Holder 3",
            //    CNICNumber = "11111-1111111-3",
            //    Department = "Department 1",
            //    Section = "Section 2",
            //    Cadre = "NMPT",
            //    NetNormalHours = 8,
            //    OverTimeHours = 2,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2
            //});

            //cnicWiseReportInfo.Add("11111-1111111-4", new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date,
            //    FirstName = "Card Holder 4",
            //    CNICNumber = "11111-1111111-4",
            //    Department = "Department 1",
            //    Section = "Section 1",
            //    Cadre = "NMPT 2",
            //    NetNormalHours = 8,
            //    OverTimeHours = 2,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2
            //});

            //cnicWiseReportInfo.Add("11111-1111111-5", new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date,
            //    FirstName = "Card Holder 5",
            //    CNICNumber = "11111-1111111-5",
            //    Department = "Department 1",
            //    Section = "Section 2",
            //    Cadre = "NMPT",
            //    NetNormalHours = 8,
            //    OverTimeHours = 2,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2
            //});

            //cnicWiseReportInfo.Add("11111-1111111-6", new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date,
            //    FirstName = "Card Holder 6",
            //    CNICNumber = "11111-1111111-6",
            //    Department = "Department 1",
            //    Section = "Section 1",
            //    Cadre = "NMPT",
            //    NetNormalHours = 8,
            //    OverTimeHours = 2,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2
            //});

            //cnicWiseReportInfo.Add("11111-1111111-7", new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date,
            //    FirstName = "Card Holder 7",
            //    CNICNumber = "11111-1111111-7",
            //    Department = "Department 2",
            //    Section = "Section 2",
            //    Cadre = "NMPT",
            //    NetNormalHours = 8,
            //    OverTimeHours = 2,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2
            //});

            //cnicWiseReportInfo.Add("11111-1111111-8", new CardHolderReportInfo()
            //{
            //    OccurrenceTime = DateTime.Now.Date,
            //    FirstName = "Card Holder 8",
            //    CNICNumber = "11111-1111111-8",
            //    Department = "Department 8",
            //    Section = "Section 8",
            //    Cadre = "NMPT",
            //    NetNormalHours = 8,
            //    OverTimeHours = 2,
            //    CallOutFrom = "18:00",
            //    CallOutTo = "20:00",
            //    TotalCallOutHours = 2
            //});


            #endregion

            if (cnicWiseReportInfo != null && cnicWiseReportInfo.Keys.Count > 0)
            {
                this.mData = new Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>>();

                foreach (KeyValuePair<string, CardHolderReportInfo> reportInfo in cnicWiseReportInfo)
                {
                    if (reportInfo.Value == null)
                    {
                        continue;
                    }

                    string department = reportInfo.Value.Department;
                    string section = reportInfo.Value.Section;
                    string cadre = reportInfo.Value.Cadre;

                    if (this.mData.ContainsKey(department))
                    {
                        if (this.mData[department].ContainsKey(section))
                        {
                            if (this.mData[department][section].ContainsKey(cadre))
                            {
                                this.mData[department][section][cadre].Add(reportInfo.Value);
                            }
                            else
                            {
                                this.mData[department][section].Add(cadre, new List<CardHolderReportInfo>() { reportInfo.Value });
                            }
                        }
                        else
                        {
                            Dictionary<string, List<CardHolderReportInfo>> cadreWiseList = new Dictionary<string, List<CardHolderReportInfo>>();
                            cadreWiseList.Add(cadre, new List<CardHolderReportInfo>() { reportInfo.Value });
                            this.mData[department].Add(section, cadreWiseList);
                        }
                    }
                    else
                    {
                        Dictionary<string, List<CardHolderReportInfo>> cadreWiseList = new Dictionary<string, List<CardHolderReportInfo>>();
                        cadreWiseList.Add(cadre, new List<CardHolderReportInfo>() { reportInfo.Value });

                        Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>> sectionWiseReport = new Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>();
                        sectionWiseReport.Add(section, cadreWiseList);

                        this.mData.Add(department, sectionWiseReport);
                    }
                }
            }
            else
            {
                MessageBox.Show(this, "No Data found.");
            }

            if (this.mData != null && this.mData.Count > 0)
            {
                //Cursor.Current = currentCursor;
                this.saveFileDialog1.ShowDialog(this);
            }
            else
            {
                //Cursor.Current = currentCursor;
                MessageBox.Show(this, "No data exist on current selected date range.");
            }
            


            //create data object and print report.


        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

            string extension = Path.GetExtension(this.saveFileDialog1.FileName);

            if (extension == ".pdf")
            {
                this.SaveAsPdf(this.mData, "Attendance Report");
            }
            else if (extension == ".xlsx")
            {
                this.SaveAsExcel(this.mData, "Attendance Report", "Attendance Report");
            }
        }

        private void SaveAsPdf(Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>> data, string heading)
        {
            Cursor currentCursor = Cursor.Current;

            try
            {
                Cursor.Current = Cursors.WaitCursor;
                if (data != null)
                {
                    using (PdfWriter pdfWriter = new PdfWriter(this.saveFileDialog1.FileName))
                    {
                        using (PdfDocument pdfDocument = new PdfDocument(pdfWriter))
                        {
                            using (Document doc = new Document(pdfDocument))
                            {
                                doc.SetFont(PdfFontFactory.CreateFont("Fonts/SEGOEUIL.TTF"));
                                string headerLeftText = "Report From: " + this.dtpFromDate.Value.ToShortDateString() + " To: " + this.dtpToDate.Value.ToShortDateString();
                                string headerRightText = string.Empty;
                                string footerLeftText = "This is computer generated report.";
                                string footerRightText = "Report generated on: " + DateTime.Now.ToString();


                                pdfDocument.AddEventHandler(PdfDocumentEvent.START_PAGE, new PdfHeaderAndFooter(doc, true, headerLeftText, headerRightText));
                                pdfDocument.AddEventHandler(PdfDocumentEvent.END_PAGE, new PdfHeaderAndFooter(doc, false, footerLeftText, footerRightText));

                                pdfDocument.SetDefaultPageSize(new iText.Kernel.Geom.PageSize(1000F, 842F));
                                Table table = new Table((new List<float>() { 8F, 90F, 190F, 70F, 70F, 80F, 70F, 90F, 70F, 70F, 70F}).ToArray());

                                table.SetWidth(880F);
                                table.SetFixedLayout();
                                //Table table = new Table((new List<float>() { 8F, 100F, 150F, 225F, 60F, 40F, 100F, 125F, 150F }).ToArray());

                                this.AddMainHeading(table, heading);

                                this.AddNewEmptyRow(table);
                                //this.AddNewEmptyRow(table);

                                //Sections and Data

                                foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>> department in data)
                                {
                                    if (department.Value == null)
                                    {
                                        continue;
                                    }

                                    //Department
                                    this.AddDepartmentRow(table, department.Key);

                                    foreach (KeyValuePair<string, Dictionary<string, List<CardHolderReportInfo>>> section in department.Value)
                                    {
                                        if (section.Value == null)
                                        {
                                            continue;
                                        }

                                        //Section
                                        this.AddSectionRow(table, section.Key);


                                        foreach (KeyValuePair<string, List<CardHolderReportInfo>> cadre in section.Value)
                                        {
                                            if (cadre.Value == null)
                                            {
                                                continue;
                                            }

                                            //Cadre
                                            this.AddCadreRow(table, cadre.Key);

                                            //Data
                                            //this.AddNewEmptyRow(table, false);

                                            this.AddTableHeaderRow(table);

                                            for (int i = 0; i < cadre.Value.Count; i++)
                                            {
                                                CardHolderReportInfo chl = cadre.Value[i];
                                                this.AddTableDataRow(table, chl, i % 2 == 0);
                                            }

                                            this.AddNewEmptyRow(table);
                                        }
                                    }
                                }



                                doc.Add(table);

                                doc.Close();
                            }
                        }

                        System.Diagnostics.Process.Start(this.saveFileDialog1.FileName);
                    }
                }
                Cursor.Current = currentCursor;
            }
            catch (Exception exp)
            {
                Cursor.Current = currentCursor;
                if (exp.HResult == -2147024864)
                {
                    MessageBox.Show(this, "\"" + this.saveFileDialog1.FileName + "\" is already is use.\n\nPlease close it and generate report again.");
                }
                else
                if (exp.HResult == -2147024891)
                {
                    MessageBox.Show(this, "You did not have rights to save file on selected location.\n\nPlease run as administrator.");
                }
                else
                {
                    MessageBox.Show(this, exp.Message);
                }

            }
        }

        private void SaveAsExcel(Dictionary<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>> data, string sheetName, string heading)
        {
            Cursor currentCursor = Cursor.Current;
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                if (data != null)
                {



                    using (ExcelPackage ex = new ExcelPackage())
                    {
                        ExcelWorksheet work = ex.Workbook.Worksheets.Add(sheetName);

                        work.View.ShowGridLines = false;
                        work.Cells.Style.Font.Name = "Segoe UI Light";

                        work.Column(1).Width = 15.14;
                        work.Column(2).Width = 37.29;
                        work.Column(3).Width = 15.14;
                        work.Column(4).Width = 15.14;
                        work.Column(5).Width = 20;
                        work.Column(6).Width = 15.14;
                        work.Column(7).Width = 20;
                        work.Column(8).Width = 15.14;
                        work.Column(9).Width = 15.14;
                        work.Column(10).Width = 15.14;

                        //Heading
                        work.Cells["A1:B2"].Merge = true;
                        work.Cells["A1:B2"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        work.Cells["A1:B2"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(252, 213, 180));
                        work.Cells["A1:B2"].Style.Font.Size = 22;
                        work.Cells["A1:B2"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells["A1:B2"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        //work.Cells["A1:B2"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        //work.Cells["A1:B2"].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        //work.Cells["A1:B2"].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        //work.Cells["A1:B2"].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        //work.Cells["A1:B2"].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells["A1:B2"].Value = heading;

                        // img variable actually is your image path
                        System.Drawing.Image myImage = System.Drawing.Image.FromFile("Images/logo.png");

                        var pic = work.Drawings.AddPicture("Logo", myImage);

                        pic.SetPosition(5, 1100);

                        int row = 4;

                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1].Value = "Report From: ";
                        work.Cells[row, 2].Value = this.dtpFromDate.Value.ToShortDateString();
                        work.Row(row).Height = 20;

                        row++;
                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1].Value = "Report To:";
                        work.Cells[row, 2].Value = this.dtpToDate.Value.ToShortDateString();
                        work.Row(row).Height = 20;

                        //row++;
                        //work.Cells[row, 1].Style.Font.Bold = true;
                        //work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        //work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        //work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        //work.Cells[row, 1].Value = "Late Time Range:";
                        //work.Cells[row, 2].Value = this.dtpLateTimeStart.Value.ToShortTimeString() + " - " + this.dtpLateTimeEnd.Value.ToShortTimeString();
                        //work.Row(row).Height = 20;

                        row++;
                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1].Value = "Report Time: ";
                        work.Cells[row, 2].Value = DateTime.Now.ToString();
                        work.Row(row).Height = 20;

                        row++;
                        row++;
                        //Sections and Data

                        foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, List<CardHolderReportInfo>>>> department in data)
                        {
                            if (department.Value == null)
                            {
                                continue;
                            }

                            //Department
                            work.Cells[row, 1].Style.Font.Bold = true;
                            work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                            work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            work.Cells[row, 1].Value = "Department:";
                            work.Cells[row, 2].Value = department.Key;
                            work.Cells[row, 2].Style.Font.UnderLine = true;
                            work.Row(row).Height = 20;

                            row++;

                            foreach (KeyValuePair<string, Dictionary<string, List<CardHolderReportInfo>>> section in department.Value)
                            {
                                if (section.Value == null)
                                {
                                    continue;
                                }

                                //Section
                                work.Cells[row, 1].Style.Font.Bold = true;
                                work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                work.Cells[row, 1].Value = "Section:";
                                work.Cells[row, 2].Value = section.Key;
                                work.Row(row).Height = 20;

                                //Data
                                row++;

                                foreach (KeyValuePair<string, List<CardHolderReportInfo>> cadre in section.Value)
                                {

                                    if (cadre.Value == null)
                                    {
                                        continue;
                                    }

                                    //Section
                                    work.Cells[row, 1].Style.Font.Bold = true;
                                    work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                    work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                    work.Cells[row, 1].Value = "Cadre:";
                                    work.Cells[row, 2].Value = cadre.Key;
                                    work.Row(row).Height = 20;

                                    //Data
                                    row++;


                                    work.Cells[row, 1, row, 10].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    work.Cells[row, 1, row, 10].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    work.Cells[row, 1, row, 10].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    work.Cells[row, 1, row, 10].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                    work.Cells[row, 1, row, 10].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 10].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 10].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 10].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                                    work.Cells[row, 1, row, 10].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    work.Cells[row, 1, row, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(253, 233, 217));
                                    work.Cells[row, 1, row, 10].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                    work.Cells[row, 1, row, 10].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                    work.Cells[row, 1].Value = "Date";
                                    work.Cells[row, 2].Value = "First Name";
                                    work.Cells[row, 3].Value = "P-Number";
                                    work.Cells[row, 4].Value = "Cadre";
                                    work.Cells[row, 5].Value = "Net Normal Hrs";
                                    work.Cells[row, 6].Value = "OT Hrs";
                                    work.Cells[row, 7].Value = "Normal + OT Hrs";
                                    work.Cells[row, 8].Value = "CO From Hrs";
                                    work.Cells[row, 9].Value = "CO To Hrs";
                                    work.Cells[row, 10].Value = "CO Total Hrs";

                                    work.Row(row).Height = 20;

                                    for (int i = 0; i < cadre.Value.Count; i++)
                                    {
                                        row++;
                                        work.Cells[row, 1, row, 10].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        work.Cells[row, 1, row, 10].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        work.Cells[row, 1, row, 10].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                        work.Cells[row, 1, row, 10].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                        work.Cells[row, 1, row, 10].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 10].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 10].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                        work.Cells[row, 1, row, 10].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                                        if (i % 2 == 0)
                                        {
                                            work.Cells[row, 1, row, 10].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                            work.Cells[row, 1, row, 10].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                        }

                                        work.Cells[row, 1, row, 10].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        //work.Cells[row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                        //work.Cells[row, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                                        CardHolderReportInfo chl = cadre.Value[i];
                                        work.Cells[row, 1].Value = chl.OccurrenceTime.Date.ToShortDateString();
                                        work.Cells[row, 2].Value = chl.FirstName;
                                        work.Cells[row, 3].Value = chl.PNumber;
                                        work.Cells[row, 4].Value = chl.Cadre;
                                        work.Cells[row, 5].Value = chl.NetNormalHours.ToString();
                                        work.Cells[row, 6].Value = chl.OverTimeHours.ToString();
                                        work.Cells[row, 7].Value = (chl.NetNormalHours + chl.OverTimeHours).ToString();
                                        work.Cells[row, 8].Value = chl.CallOutFrom;
                                        work.Cells[row, 9].Value = chl.CallOutTo;
                                        work.Cells[row, 10].Value = chl.TotalCallOutHours.ToString();


                                        work.Row(row).Height = 20;
                                    }

                                    row++;
                                    row++;
                                }


                            }

                        }

                        ex.SaveAs(new System.IO.FileInfo(this.saveFileDialog1.FileName));

                        System.Diagnostics.Process.Start(this.saveFileDialog1.FileName);
                    }
                }
                Cursor.Current = currentCursor;
            }
            catch (Exception exp)
            {
                Cursor.Current = currentCursor;
                if (exp.InnerException != null && exp.InnerException.InnerException != null)
                {
                    if (exp.InnerException.InnerException.HResult == -2147024864)
                    {
                        MessageBox.Show(this, "\"" + this.saveFileDialog1.FileName + "\" is already is use.\n\nPlease close it and generate report again.");
                    }
                    if (exp.InnerException.InnerException.HResult == -2147024891)
                    {
                        MessageBox.Show(this, "You did not have rights to save file on selected location.\n\nPlease run as administrator.");
                    }
                }
                else
                {
                    MessageBox.Show(this, exp.Message);
                }

            }

        }

        private void AddMainHeading(Table table, string heading)
        {
            Cell headingCell = new Cell(2, 6);
            headingCell.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
            headingCell.SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3));
            headingCell.Add(new Paragraph(heading).SetFontSize(22F).SetBackgroundColor(new DeviceRgb(252, 213, 180))
                // .SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 3))
                );
            iText.Layout.Element.Image img = new iText.Layout.Element.Image(iText.IO.Image.ImageDataFactory.Create("Images/logo.png"));

            table.AddCell(headingCell);
            table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            //table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            //table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            table.AddCell(new Cell().Add(img).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
        }

        private void AddNewEmptyRow(Table table, bool removeBottomBorder = true)
        {
            table.StartNewRow();

            if (removeBottomBorder)
            {
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(6F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            }
            else
            {
                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(6F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            }

        }

        private void AddDepartmentRow(Table table, string departmentName)
        {
            table.StartNewRow();

            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph("Department:").
                    SetFontSize(11F).
                    SetBold().
                    SetFontColor(new DeviceRgb(247, 150, 70))).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph(departmentName).
                    SetFontSize(11F).SetUnderline()).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
        }

        private void AddSectionRow(Table table, string sectionName)
        {
            table.StartNewRow();

            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph("Section:").
                    SetFontSize(11F).
                    SetBold().
                    SetFontColor(new DeviceRgb(247, 150, 70))).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph(sectionName).
                    SetFontSize(11F)).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
        }

        private void AddCadreRow(Table table, string cadreName)
        {
            table.StartNewRow();

            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph("Cadre:").
                    SetFontSize(11F).
                    SetBold().
                    SetFontColor(new DeviceRgb(247, 150, 70))).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph(cadreName).
                    SetFontSize(11F)).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().SetHeight(22F).SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
        }

        private void AddTableHeaderRow(Table table)
        {
            table.StartNewRow();

            table.AddCell(new Cell().
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderBottom(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph("Date").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("First Name").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("P-Number").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Cadre").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Net Normal Hrs").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("OT Hrs").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Normal + OT Hrs").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("CO From Hrs").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("CO To Hrs").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("CO Total Hrs").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
        }

        private void AddTableDataRow(Table table, CardHolderReportInfo chl, bool altRow)
        {
            if (chl == null)
            {
                return;
            }



            table.StartNewRow();

            table.AddCell(new Cell().
                SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                SetBorderBottom(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph(chl.OccurrenceTime.Date.ToShortDateString()).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.FirstName) ? string.Empty : chl.FirstName).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.PNumber) ? string.Empty : chl.PNumber).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.Cadre) ? string.Empty : chl.Cadre).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(chl.NetNormalHours.ToString()).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(chl.OverTimeHours.ToString()).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph((chl.NetNormalHours + chl.OverTimeHours).ToString()).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.CallOutFrom) ? string.Empty : chl.CallOutFrom).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.CallOutTo) ? string.Empty : chl.CallOutTo).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(chl.TotalCallOutHours.ToString()).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
        }
    }
}
