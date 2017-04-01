using AttendanceReport.CCFTCentral;
using AttendanceReport.CCFTEvent;
using iText.IO.Font;
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
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AttendanceReport
{
    public partial class Form1 : Form
    {
        public Dictionary<string, Dictionary<string, List<CardHolderInfo>>> mData = null;
        public Task<List<string>> mDepartmentTask = null;
        public Task<List<string>> mSectionTask = null;
        public Task<List<string>> mCrewTask = null;
        public Task<List<string>> mCadreTask = null;
        public Task<List<string>> mCompanyTask = null;

        public Form1()
        {
            InitializeComponent();



            this.mDepartmentTask = new Task<List<string>>(() =>
            {
                CCFTCentral.CCFTCentral ccftCentral = new CCFTCentral.CCFTCentral();

                List<string> departments = (from pds in ccftCentral.PersonalDataStrings
                                            where pds != null && pds.PersonalDataFieldID == 5043 && pds.Value != null
                                            select pds.Value).Distinct().ToList();

                return departments;
            });

            this.mSectionTask = new Task<List<string>>(() =>
            {
                CCFTCentral.CCFTCentral ccftCentral = new CCFTCentral.CCFTCentral();

                List<string> sections = (from pds in ccftCentral.PersonalDataStrings
                                         where pds != null && pds.PersonalDataFieldID == 12951 && pds.Value != null
                                         select pds.Value).Distinct().ToList();

                return sections;
            });

            this.mCrewTask = new Task<List<string>>(() =>
            {
                CCFTCentral.CCFTCentral ccftCentral = new CCFTCentral.CCFTCentral();

                List<string> crews = (from pds in ccftCentral.PersonalDataStrings
                                      where pds != null && pds.PersonalDataFieldID == 12869 && pds.Value != null
                                      select pds.Value).Distinct().ToList();

                return crews;
            });

            this.mCadreTask = new Task<List<string>>(() =>
            {
                CCFTCentral.CCFTCentral ccftCentral = new CCFTCentral.CCFTCentral();

                List<string> cadres = (from pds in ccftCentral.PersonalDataStrings
                                            where pds != null && pds.PersonalDataFieldID == 12952 && pds.Value != null
                                            select pds.Value).Distinct().ToList();

                return cadres;
            });

            this.mCompanyTask = new Task<List<string>>(() =>
            {
                CCFTCentral.CCFTCentral ccftCentral = new CCFTCentral.CCFTCentral();

                List<string> companyNames = (from pds in ccftCentral.PersonalDataStrings
                                       where pds != null && pds.PersonalDataFieldID == 5059 && pds.Value != null
                                       select pds.Value).Distinct().ToList();

                return companyNames;
            });

            this.mDepartmentTask.Start();
            this.mSectionTask.Start();
            this.mCrewTask.Start();
            this.mCadreTask.Start();
            this.mCompanyTask.Start();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Cursor currentCursor = Cursor.Current;
            try
            {

                Cursor.Current = Cursors.WaitCursor;
                this.mData = null;

                DateTime fromDate = this.dtpFromDate.Value.Date.ToUniversalTime();
                DateTime toDate = this.dtpToDate.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59).ToUniversalTime();




                CCFTEvent.CCFTEvent ccftEvent = new CCFTEvent.CCFTEvent();

                CCFTCentral.CCFTCentral ccftCentral = new CCFTCentral.CCFTCentral();
                List<int> ids = new List<int>();



                #region Events

                List<CCFTEvent.Event> lstEvents = (from events in ccftEvent.Events
                                         where
                                             events != null && events.EventType == 20001 &&
                                             events.OccurrenceTime >= fromDate &&
                                             events.OccurrenceTime < toDate
                                         select events).ToList();



                Dictionary<int, DateTime> lstChlEvents = new Dictionary<int, DateTime>();

                foreach (CCFTEvent.Event events in lstEvents)
                {
                    if (events == null || events.RelatedItems == null)
                    {
                        continue;
                    }

                    foreach (RelatedItem relatedItem in events.RelatedItems)
                    {
                        if (relatedItem != null && relatedItem.RelationCode == 0)
                        {
                            ids.Add(relatedItem.FTItemID);
                            if (lstChlEvents.ContainsKey(relatedItem.FTItemID))
                            {
                                DateTime occurranceTime = lstChlEvents[relatedItem.FTItemID];

                                if (occurranceTime > events.OccurrenceTime)
                                {
                                    lstChlEvents[relatedItem.FTItemID] = events.OccurrenceTime;
                                }

                            }
                            else
                            {
                                lstChlEvents.Add(relatedItem.FTItemID, events.OccurrenceTime);
                            }

                        }
                    }
                }
                #endregion

                #region CHL
                //List<CardholderLocation> cardHolderLocations = (from cardHolder in ccftCentral.CardholderLocations
                //                                                where
                //                                                   cardHolder != null &&
                //                                                   cardHolder.AccessTime >= fromDate &&
                //                                                   cardHolder.AccessTime < toDate &&
                //                                                   cardHolder.AccessType == 1
                //                                                select cardHolder).ToList();


                //ids = (from chl in cardHolderLocations
                //                 where chl != null
                //                 select chl.CardholderID).ToList();
                #endregion

                List<Card> cards = (from card in ccftCentral.Cards
                                    where card != null && ids.Contains(card.CardholderID)
                                    select card).Distinct().ToList();

                #region CHL
                //List<Cardholder> cardHolders = (from card in cards
                //                                where card != null
                //                                select card.Cardholder).ToList();

                //Dictionary<int, CardholderLocation> filteredChls =
                //                            new Dictionary<int, CardholderLocation>();

                //foreach (CardholderLocation chl in cardHolderLocations)
                //{
                //    if (filteredChls.ContainsKey(chl.CardholderID))
                //    {
                //        CardholderLocation chlExist = filteredChls[chl.CardholderID];

                //        if (chlExist.AccessTime > chl.AccessTime)
                //        {
                //            filteredChls[chl.CardholderID] = chl;
                //        }

                //    }
                //    else
                //    {
                //        filteredChls.Add(chl.CardholderID, chl);
                //    }
                //}
                #endregion

                this.mData = new Dictionary<string, Dictionary<string, List<CardHolderInfo>>>();
                string filterByDepartment = this.cbxDepartments.Text;
                string filterBySection = this.cbxSections.Text;
                string filerByName = this.tbxName.Text;
                string filterByPNumber = this.tbxPNumber.Text;
                string filterByCardNumber = this.tbxCarNumber.Text;
                string filterByCrew = this.cbxCrew.Text;
                string filterByCadre = this.cbxCadre.Text;
                string filterByCompany = this.cbxCompany.Text;
                string filterByCNIC = this.tbxCnic.Text;

                #region CHL
                //foreach (KeyValuePair<int, CardholderLocation> chlEvent in filteredChls)
                #endregion
                #region Events
                foreach (KeyValuePair<int, DateTime> chlEvent in lstChlEvents)
                #endregion
                {
                    if (cards.Exists(c => c.CardholderID == chlEvent.Key))
                    {
                        Card card = cards.Find(c => c.CardholderID == chlEvent.Key);
                        #region CHL
                        //CardholderLocation cardHolderLocation = chlEvent.Value;
                        #endregion
                        #region Events
                        DateTime occurranceTime = chlEvent.Value.AddHours(5);
                        #endregion
                        Cardholder cardHolder = card.Cardholder;

                        if (cardHolder != null && cardHolder.PersonalDataStrings != null)
                        {
                            string department = (from pds in cardHolder.PersonalDataStrings
                                                 where pds != null && pds.PersonalDataFieldID == 5043
                                                 select pds.Value).FirstOrDefault();

                            string section = (from pds in cardHolder.PersonalDataStrings
                                              where pds != null && pds.PersonalDataFieldID == 12951
                                              select pds.Value).FirstOrDefault();

                            string crew = (from pds in cardHolder.PersonalDataStrings
                                           where pds != null && pds.PersonalDataFieldID == 12869
                                           select pds.Value).FirstOrDefault();

                            string cadre = (from pds in cardHolder.PersonalDataStrings
                                              where pds != null && pds.PersonalDataFieldID == 12952
                                              select pds.Value).FirstOrDefault();

                            string cnic = (from pds in cardHolder.PersonalDataStrings
                                            where pds != null && pds.PersonalDataFieldID == 5051
                                            select pds.Value).FirstOrDefault();

                            string company = (from pds in cardHolder.PersonalDataStrings
                                            where pds != null && pds.PersonalDataFieldID == 5059
                                            select pds.Value).FirstOrDefault();

                            
                            if (!string.IsNullOrEmpty(department))
                            {
                                department = department.ToUpper();
                                filterByDepartment = filterByDepartment.ToUpper();


                                //Filter By Department
                                if (!string.IsNullOrEmpty(filterByDepartment) && department != filterByDepartment)
                                {
                                    continue;
                                }

                                if (!string.IsNullOrEmpty(section))
                                {
                                    section = section.ToUpper();
                                    filterBySection = filterBySection.ToUpper();


                                    //Filter By Section
                                    if (!string.IsNullOrEmpty(filterBySection) && section != filterBySection)
                                    {
                                        continue;
                                    }


                                    //Filter By Name
                                    if (!string.IsNullOrEmpty(filerByName) && !cardHolder.FirstName.Contains(filerByName))
                                    {
                                        continue;
                                    }


                                    //Filter By P-Number
                                    int? pNumber = cardHolder.PersonalDataIntegers == null || cardHolder.PersonalDataIntegers.Count == 0 ? null : cardHolder.PersonalDataIntegers.ElementAt(0).Value;
                                    string strPNumber = pNumber == null ? "Nil" : pNumber.ToString();

                                    if (!string.IsNullOrEmpty(filterByPNumber) && (pNumber == null || pNumber != Convert.ToInt32(filterByPNumber)))
                                    {
                                        continue;
                                    }


                                    //Filter By Card Number
                                    if (!string.IsNullOrEmpty(filterByCardNumber))
                                    {
                                        int cardNumber;

                                        bool parsed = Int32.TryParse(cardHolder.LastName, out cardNumber);

                                        if (parsed)
                                        {
                                            if (cardNumber != Convert.ToInt32(filterByCardNumber))
                                            {
                                                continue;
                                            }
                                        }
                                        else
                                        {
                                            continue;
                                        }

                                    }

                                    //Filter By Crew
                                    if (!string.IsNullOrEmpty(filterByCrew) && crew != filterByCrew)
                                    {
                                        continue;
                                    }
                                    
                                    //Filter By Cadre
                                    if (!string.IsNullOrEmpty(filterByCadre) && cadre != filterByCadre)
                                    {
                                        continue;
                                    }

                                    //Filter By CNIC
                                    if (!string.IsNullOrEmpty(filterByCNIC) && cnic != filterByCNIC)
                                    {
                                        continue;
                                    }
                                    
                                    //Filter By Company
                                    if (!string.IsNullOrEmpty(filterByCompany) && company != filterByCompany)
                                    {
                                        continue;
                                    }

                                    TimeSpan thStartTime = this.dtpLateTimeStart.Value.TimeOfDay;
                                    TimeSpan thEndTime = this.dtpLateTimeEnd.Value.TimeOfDay;

                                    #region CHL
                                    //if (TimeSpan.Compare(cardHolderLocation.AccessTime.AddHours(5).TimeOfDay, thStartTime) > 0 && TimeSpan.Compare(cardHolderLocation.AccessTime.AddHours(5).TimeOfDay, thEndTime) <= 0)
                                    #endregion
                                    #region Events
                                    if (TimeSpan.Compare(occurranceTime.TimeOfDay, thStartTime) > 0 && TimeSpan.Compare(occurranceTime.TimeOfDay, thEndTime) <= 0)
                                    #endregion
                                    {
                                        CardHolderInfo chi = new CardHolderInfo()
                                        {
                                            CardNumber = cardHolder.LastName,
                                            FirstName = cardHolder.FirstName,
                                            #region CHL
                                            //OccurrenceTime = cardHolderLocation.AccessTime.AddHours(5),
                                            #endregion
                                            #region Events
                                            OccurrenceTime = occurranceTime,
                                            #endregion
                                            PNumber = strPNumber,
                                            Crew = crew,
                                            Cadre = cadre,
                                            Company = company,
                                            CNICNumber = cnic
                                        };

                                        if (this.mData.ContainsKey(department))
                                        {
                                            if (this.mData[department].ContainsKey(section))
                                            {
                                                List<CardHolderInfo> lstchi = this.mData[department][section];
                                                lstchi.Add(chi);
                                                this.mData[department][section] = lstchi;
                                            }
                                            else
                                            {
                                                List<CardHolderInfo> lstchi = new List<CardHolderInfo>();
                                                lstchi.Add(chi);
                                                this.mData[department].Add(section, lstchi);
                                            }
                                        }
                                        else
                                        {
                                            List<CardHolderInfo> lstchi = new List<CardHolderInfo>();
                                            lstchi.Add(chi);
                                            Dictionary<string, List<CardHolderInfo>> dict = new Dictionary<string, List<CardHolderInfo>>();
                                            dict.Add(section, lstchi);
                                            this.mData.Add(department, dict);
                                        }

                                    }
                                }
                            }
                        }
                    }
                }

                if (this.mData != null && this.mData.Count > 0)
                {
                    Cursor.Current = currentCursor;
                    this.saveFileDialog1.ShowDialog(this);
                }
                else
                {
                    Cursor.Current = currentCursor;
                    MessageBox.Show(this, "No data exist on current selected date range.");
                }


            }
            catch (Exception exp)
            {
                string exMessage = exp.Message;
                Exception innerException = exp.InnerException;

                while (innerException != null)
                {
                    exMessage = "\n" + innerException.Message;
                    innerException = innerException.InnerException;
                }

                Cursor.Current = currentCursor;
                MessageBox.Show(this, exMessage);
            }

        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

            string extension = Path.GetExtension(this.saveFileDialog1.FileName);

            if (extension == ".pdf")
            {
                this.SaveAsPdf();
            }
            else if (extension == ".xlsx")
            {
                this.SaveAsExcel();
            }
        }

        private void SaveAsPdf()
        {
            Cursor currentCursor = Cursor.Current;

            try
            {
                Cursor.Current = Cursors.WaitCursor;
                if (this.mData != null)
                {
                    using (PdfWriter pdfWriter = new PdfWriter(this.saveFileDialog1.FileName))
                    {
                        using (PdfDocument pdfDocument = new PdfDocument(pdfWriter))
                        {
                            using (Document doc = new Document(pdfDocument))
                            {
                                doc.SetFont(PdfFontFactory.CreateFont("Fonts/SEGOEUIL.TTF"));
                                string headerLeftText = "Report From: " + this.dtpFromDate.Value.ToShortDateString() + " To: " + this.dtpToDate.Value.ToShortDateString();
                                string headerRightText = "Late Time Range: " + this.dtpLateTimeStart.Value.ToShortTimeString() + " - " + this.dtpLateTimeEnd.Value.ToShortTimeString();
                                string footerLeftText = "This is computer generated report.";
                                string footerRightText = "Report generated on: " + DateTime.Now.ToString(); 

                                pdfDocument.AddEventHandler(PdfDocumentEvent.START_PAGE, new PdfHeaderAndFooter(doc, true, headerLeftText, headerRightText));
                                pdfDocument.AddEventHandler(PdfDocumentEvent.END_PAGE, new PdfHeaderAndFooter(doc, false, footerLeftText, footerRightText));

                                //pdfDocument.SetDefaultPageSize(new iText.Kernel.Geom.PageSize(1000F, 1000F));
                                Table table = new Table((new List<float>() { 8F, 100F, 150F, 70F, 250F }).ToArray());


                                //Table table = new Table((new List<float>() { 8F, 100F, 150F, 225F, 60F, 40F, 100F, 125F, 150F }).ToArray());

                                this.AddMainHeading(table);

                                //this.AddNewEmptyRow(table);
                                //this.AddNewEmptyRow(table);

                                //Sections and Data

                                foreach (KeyValuePair<string, Dictionary<string, List<CardHolderInfo>>> department in this.mData)
                                {
                                    if (department.Value == null)
                                    {
                                        continue;
                                    }

                                    //Department
                                    this.AddDepartmentRow(table, department.Key);


                                    foreach (KeyValuePair<string, List<CardHolderInfo>> section in department.Value)
                                    {
                                        //Section
                                        this.AddSectionRow(table, section.Key);

                                        //Data
                                        //this.AddNewEmptyRow(table, false);

                                        this.AddTableHeaderRow(table);

                                        for (int i = 0; i < section.Value.Count; i++)
                                        {
                                            CardHolderInfo chl = section.Value[i];
                                            this.AddTableDataRow(table, chl, i % 2 == 0);
                                        }

                                        this.AddNewEmptyRow(table);
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

        private void SaveAsExcel()
        {
            Cursor currentCursor = Cursor.Current;
            try
            {
                Cursor.Current = Cursors.WaitCursor;
                if (this.mData != null)
                {



                    using (ExcelPackage ex = new ExcelPackage())
                    {
                        ExcelWorksheet work = ex.Workbook.Worksheets.Add("Attendence Report");
                        
                        work.View.ShowGridLines = false;
                        work.Cells.Style.Font.Name = "Segoe UI Light";

                        work.Column(1).Width = 18.14;
                        work.Column(2).Width = 25.29;
                        work.Column(3).Width = 15.14;
                        work.Column(4).Width = 54;

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
                        work.Cells["A1:B2"].Value = "Late Arrival Report";

                        // img variable actually is your image path
                        System.Drawing.Image myImage = System.Drawing.Image.FromFile("Images/logo.png");

                        var pic = work.Drawings.AddPicture("Logo", myImage);
                        
                        pic.SetPosition(5, 600);

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

                        row++;
                        work.Cells[row, 1].Style.Font.Bold = true;
                        work.Cells[row, 1].Style.Font.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                        work.Cells[row, 1, row, 2].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells[row, 1, row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        work.Cells[row, 1].Value = "Late Time Range:";
                        work.Cells[row, 2].Value = this.dtpLateTimeStart.Value.ToShortTimeString() + " - " + this.dtpLateTimeEnd.Value.ToShortTimeString();
                        work.Row(row).Height = 20;

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

                        foreach (KeyValuePair<string, Dictionary<string, List<CardHolderInfo>>> department in this.mData)
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

                            foreach (KeyValuePair<string, List<CardHolderInfo>> section in department.Value)
                            {
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

                                work.Cells[row, 1, row, 4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                work.Cells[row, 1, row, 4].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                work.Cells[row, 1, row, 4].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                work.Cells[row, 1, row, 4].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                work.Cells[row, 1, row, 4].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                work.Cells[row, 1, row, 4].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                work.Cells[row, 1, row, 4].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                work.Cells[row, 1, row, 4].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                                work.Cells[row, 1, row, 4].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                work.Cells[row, 1, row, 4].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(253, 233, 217));
                                work.Cells[row, 1, row, 4].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                work.Cells[row, 1, row, 4].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                work.Cells[row, 1].Value = "Card Number";
                                work.Cells[row, 2].Value = "Occurrance Time";
                                work.Cells[row, 3].Value = "P-Number";
                                work.Cells[row, 4].Value = "First Name";
                                work.Row(row).Height = 20;

                                for (int i = 0; i < section.Value.Count; i++)
                                {
                                    row++;
                                    work.Cells[row, 1, row, 4].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    work.Cells[row, 1, row, 4].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    work.Cells[row, 1, row, 4].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    work.Cells[row, 1, row, 4].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                    work.Cells[row, 1, row, 4].Style.Border.Top.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 4].Style.Border.Bottom.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 4].Style.Border.Left.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 1, row, 4].Style.Border.Right.Color.SetColor(System.Drawing.Color.FromArgb(247, 150, 70));

                                    if (i % 2 == 0)
                                    {
                                        work.Cells[row, 1, row, 4].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        work.Cells[row, 1, row, 4].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                    }

                                    work.Cells[row, 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                                    work.Cells[row, 2].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                                    work.Cells[row, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;

                                    CardHolderInfo chl = section.Value[i];
                                    work.Cells[row, 1].Value = chl.CardNumber;
                                    work.Cells[row, 2].Value = chl.OccurrenceTime.ToString();
                                    work.Cells[row, 3].Value = chl.PNumber;
                                    work.Cells[row, 4].Value = chl.FirstName;
                                    
                                    work.Row(row).Height = 20;
                                }

                                row++;
                                row++;
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

        private void AddMainHeading(Table table)
        {
            Cell headingCell = new Cell(2, 4);
            headingCell.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
            headingCell.SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3));
            headingCell.Add(new Paragraph("Late Arrival Report").SetFontSize(22F).SetBackgroundColor(new DeviceRgb(252, 213, 180))
               // .SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 3))
                );
            iText.Layout.Element.Image img = new iText.Layout.Element.Image(iText.IO.Image.ImageDataFactory.Create("Images/logo.png"));

            table.AddCell(headingCell);
            table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            //table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            table.AddCell(new Cell().Add(img).SetMarginLeft(60F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
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
                //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
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

                //table.AddCell(new Cell().
                //    SetHeight(22F).
                //    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                //table.AddCell(new Cell().
                //    SetHeight(22F).
                //    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                //table.AddCell(new Cell().
                //    SetHeight(22F).
                //    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                //table.AddCell(new Cell().
                //    SetHeight(22F).
                //    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                //    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
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
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            //table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
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
                    Add(new Paragraph("Card Number").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Occurrance Time").
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
                    Add(new Paragraph("First Name").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph("Crew").
            //        SetFontSize(11F)).
            //    SetBackgroundColor(new DeviceRgb(253, 233, 217)).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph("Cadre").
            //        SetFontSize(11F)).
            //    SetBackgroundColor(new DeviceRgb(253, 233, 217)).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph("CNIC Number").
            //        SetFontSize(11F)).
            //    SetBackgroundColor(new DeviceRgb(253, 233, 217)).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph("Company Name").
            //        SetFontSize(11F)).
            //    SetBackgroundColor(new DeviceRgb(253, 233, 217)).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
        }

        private void AddTableDataRow(Table table, CardHolderInfo chl, bool altRow)
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
                    Add(new Paragraph(string.IsNullOrEmpty(chl.CardNumber) ? string.Empty : chl.CardNumber).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.OccurrenceTime.ToString()) ? string.Empty : chl.OccurrenceTime.ToString()).
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
                SetTextAlignment(iText.Layout.Properties.TextAlignment.RIGHT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.FirstName) ? string.Empty : chl.FirstName).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));            
            //table.AddCell(new Cell().
            //        Add(new Paragraph(string.IsNullOrEmpty(chl.Crew) ? string.Empty : chl.Crew).
            //        SetFontSize(11F)).
            //    SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph(string.IsNullOrEmpty(chl.Cadre) ? string.Empty : chl.Cadre).
            //        SetFontSize(11F)).
            //    SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph(string.IsNullOrEmpty(chl.CNICNumber) ? string.Empty : chl.CNICNumber).
            //        SetFontSize(11F)).
            //    SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            //table.AddCell(new Cell().
            //        Add(new Paragraph(string.IsNullOrEmpty(chl.Company) ? string.Empty : chl.Company).
            //        SetFontSize(11F)).
            //    SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
            //    SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
            //    SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
            //    SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
        }

        private void TextBox1_KeyPress(object sender, System.Windows.Forms.KeyPressEventArgs e)
        {
            if (!(Char.IsDigit(e.KeyChar) || (e.KeyChar == (char)Keys.Back)))
                e.Handled = true;
        }


        private void CbxCrew_DropDown(object sender, System.EventArgs e)
        {
            try
            {
                if (this.cbxCrew.Items == null || this.cbxCrew.Items.Count == 0)
                {
                    Cursor currentCursor = Cursor.Current;
                    Cursor.Current = Cursors.WaitCursor;

                    List<string> crews = new List<string>() { string.Empty };

                    if (this.mCrewTask == null)
                    {
                        CCFTCentral.CCFTCentral ccftCentral = new CCFTCentral.CCFTCentral();

                        crews = (from pds in ccftCentral.PersonalDataStrings
                                 where pds != null && pds.PersonalDataFieldID == 12869 && pds.Value != null
                                 select pds.Value).Distinct().ToList();
                    }
                    else
                    {
                        crews.AddRange(this.mCrewTask.Result);
                    }

                    this.cbxCrew.Items.AddRange(crews.ToArray());
                    Cursor.Current = currentCursor;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
            }            
        }

        private void CbxSections_DropDown(object sender, System.EventArgs e)
        {
            try
            {
                if (this.cbxSections.Items == null || this.cbxSections.Items.Count == 0)
                {
                    Cursor currentCursor = Cursor.Current;
                    Cursor.Current = Cursors.WaitCursor;

                    List<string> sections = new List<string>() { string.Empty };

                    if (this.mSectionTask == null)
                    {
                        CCFTCentral.CCFTCentral ccftCentral = new CCFTCentral.CCFTCentral();

                        sections = (from pds in ccftCentral.PersonalDataStrings
                                    where pds != null && pds.PersonalDataFieldID == 12951 && pds.Value != null
                                    select pds.Value).Distinct().ToList();
                    }
                    else
                    {
                        sections.AddRange(this.mSectionTask.Result);
                    }

                    this.cbxSections.Items.AddRange(sections.ToArray());
                    Cursor.Current = currentCursor;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
            }
        }

        private void CbxDepartments_DropDown(object sender, System.EventArgs e)
        {
            try
            {
                if (this.cbxDepartments.Items == null || this.cbxDepartments.Items.Count == 0)
                {
                    Cursor currentCursor = Cursor.Current;
                    Cursor.Current = Cursors.WaitCursor;

                    List<string> departments = new List<string>() { string.Empty };

                    if (this.mDepartmentTask == null)
                    {
                        CCFTCentral.CCFTCentral ccftCentral = new CCFTCentral.CCFTCentral();

                        departments = (from pds in ccftCentral.PersonalDataStrings
                                       where pds != null && pds.PersonalDataFieldID == 5043 && pds.Value != null
                                       select pds.Value).Distinct().ToList();
                    }
                    else
                    {
                        departments.AddRange(this.mDepartmentTask.Result);
                    }

                    this.cbxDepartments.Items.AddRange(departments.ToArray());
                    Cursor.Current = currentCursor;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
            }
        }

        private void cbxCadre_DropDown(object sender, EventArgs e)
        {
            try
            {
                if (this.cbxCadre.Items == null || this.cbxCadre.Items.Count == 0)
                {
                    Cursor currentCursor = Cursor.Current;
                    Cursor.Current = Cursors.WaitCursor;

                    List<string> cadres = new List<string>() { string.Empty };

                    if (this.mCadreTask == null)
                    {
                        CCFTCentral.CCFTCentral ccftCentral = new CCFTCentral.CCFTCentral();

                        cadres = (from pds in ccftCentral.PersonalDataStrings
                                       where pds != null && pds.PersonalDataFieldID == 12952 && pds.Value != null
                                       select pds.Value).Distinct().ToList();
                    }
                    else
                    {
                        cadres.AddRange(this.mCadreTask.Result);
                    }

                    this.cbxCadre.Items.AddRange(cadres.ToArray());
                    Cursor.Current = currentCursor;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
            }
        }

        private void cbxCompany_DropDown(object sender, EventArgs e)
        {
            try
            {
                if (this.cbxCompany.Items == null || this.cbxCompany.Items.Count == 0)
                {
                    Cursor currentCursor = Cursor.Current;
                    Cursor.Current = Cursors.WaitCursor;

                    List<string> companyNames = new List<string>() { string.Empty };

                    if (this.mCompanyTask == null)
                    {
                        CCFTCentral.CCFTCentral ccftCentral = new CCFTCentral.CCFTCentral();

                        companyNames = (from pds in ccftCentral.PersonalDataStrings
                                  where pds != null && pds.PersonalDataFieldID == 5059 && pds.Value != null
                                  select pds.Value).Distinct().ToList();
                    }
                    else
                    {
                        companyNames.AddRange(this.mCompanyTask.Result);
                    }

                    this.cbxCompany.Items.AddRange(companyNames.ToArray());
                    Cursor.Current = currentCursor;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
            }
        }
    }

    public class PdfHeaderAndFooter : IEventHandler
    {
        public Document mDoc = null;
        bool mStart = false;
        string mLeftText, mRightText;

        public PdfHeaderAndFooter(Document doc, bool start, string leftText, string rightText)
        {
            this.mDoc = doc;
            this.mStart = start;
            this.mLeftText = leftText;
            this.mRightText = rightText;
        }

        public void HandleEvent(iText.Kernel.Events.Event @event)
        {
            PdfDocumentEvent docEvent = (PdfDocumentEvent)@event;
            PdfCanvas canvas = new PdfCanvas(docEvent.GetPage());
            iText.Kernel.Geom.Rectangle pageSize = docEvent.GetPage().GetPageSize();
            canvas.BeginText();

            try
            {
                canvas.SetFontAndSize(PdfFontFactory.CreateFont("Fonts/SEGOEUIL.TTF"), 8);
            }
            catch (IOException e)
            {
            }

            float height = pageSize.GetHeight();
            float width = pageSize.GetWidth();
            float left = pageSize.GetLeft();
            float right = pageSize.GetRight();
            float leftMargin = this.mDoc.GetLeftMargin();
            float rightMargin = this.mDoc.GetRightMargin();
            float topMargin = this.mDoc.GetTopMargin();
            float bottomMargin = this.mDoc.GetBottomMargin();
            float top = pageSize.GetTop();

            if (mStart)
            {
                canvas.SetStrokeColor(new DeviceRgb(247, 150, 70)).MoveTo(leftMargin, height - topMargin+4).LineTo(width - rightMargin, height - topMargin+4).Stroke().SetStrokeColor(iText.Kernel.Colors.Color.BLACK).
                    MoveText(leftMargin, height - topMargin + 14).ShowText(this.mLeftText).
                    MoveText(390, 0).ShowText(this.mRightText).
                   EndText().
                   Release();
            }
            else
            {
                canvas.SetStrokeColor(new DeviceRgb(247, 150, 70)).MoveTo(leftMargin, bottomMargin-2).LineTo(width - rightMargin, bottomMargin-2).Stroke().SetStrokeColor(iText.Kernel.Colors.Color.BLACK).
                    MoveText(leftMargin, bottomMargin - 12).ShowText(this.mLeftText).
                    MoveText(width - rightMargin - 180, 0).ShowText(this.mRightText).
                   EndText().
                   Release();
            }            
        }
    }
}
