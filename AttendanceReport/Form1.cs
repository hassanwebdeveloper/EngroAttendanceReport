using AttendanceReport.CCFTCentral;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
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

                DateTime fromDate = this.dtpFromDate.Value.Date;
                DateTime toDate = this.dtpToDate.Value.Date.AddHours(23).AddMinutes(59).AddSeconds(59);
                




                CCFTCentral.CCFTCentral ccftCentral = new CCFTCentral.CCFTCentral();

                List<CardholderLocation> cardHolderLocations = (from cardHolder in ccftCentral.CardholderLocations
                                                                where
                                                                   cardHolder != null &&
                                                                   cardHolder.AccessTime >= fromDate &&
                                                                   cardHolder.AccessTime < toDate &&
                                                                   cardHolder.AccessType == 1
                                                                select cardHolder).ToList();

                List<int> ids = (from chl in cardHolderLocations
                                 where chl != null
                                 select chl.CardholderID).ToList();

                List<Card> cards = (from card in ccftCentral.Cards
                                    where card != null && ids.Contains(card.CardholderID)
                                    select card).Distinct().ToList();

                List<Cardholder> cardHolders = (from card in cards
                                                where card != null
                                                select card.Cardholder).ToList();

                Dictionary<int, CardholderLocation> filteredChls =
                                            new Dictionary<int, CardholderLocation>();

                foreach (CardholderLocation chl in cardHolderLocations)
                {
                    if (filteredChls.ContainsKey(chl.CardholderID))
                    {
                        CardholderLocation chlExist = filteredChls[chl.CardholderID];

                        if (chlExist.AccessTime > chl.AccessTime)
                        {
                            filteredChls[chl.CardholderID] = chl;
                        }

                    }
                    else
                    {
                        filteredChls.Add(chl.CardholderID, chl);
                    }
                }

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

                foreach (KeyValuePair<int, CardholderLocation> chl in filteredChls)
                {
                    if (cards.Exists(c => c.CardholderID == chl.Key))
                    {
                        Card card = cards.Find(c => c.CardholderID == chl.Key);
                        CardholderLocation cardHolderLocation = chl.Value;
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

                            //for (int i = 0; i < lstPds.Count; i++)
                            //{
                            //    if (lstPds[i] != null && lstPds[i].PersonalDataFieldID == 12951)
                            //    {
                            //        section = lstPds[i].Value;
                            //        break;
                            //    }
                            //}
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

                                    if (TimeSpan.Compare(cardHolderLocation.AccessTime.TimeOfDay, thStartTime) > 0 && TimeSpan.Compare(cardHolderLocation.AccessTime.TimeOfDay, thEndTime) <= 0)
                                    {
                                        CardHolderInfo chi = new CardHolderInfo()
                                        {
                                            CardNumber = cardHolder.LastName,
                                            FirstName = cardHolder.FirstName,
                                            OccurrenceTime = cardHolderLocation.AccessTime,
                                            //TimeSpan.Compare(cardHolderLocation.AccessTime.TimeOfDay, thTime) > 0 ? "Late" : "On Time"
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
                Cursor.Current = currentCursor;
                MessageBox.Show(this, exp.Message);
            }

        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
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
                                pdfDocument.SetDefaultPageSize(new iText.Kernel.Geom.PageSize(1000F, 1000F));
                                

                                Table table = new Table((new List<float>() { 8F, 100F, 150F, 225F, 60F, 40F, 100F, 125F, 150F }).ToArray());

                                this.AddMainHeading(table);

                                this.AddNewEmptyRow(table);
                                this.AddNewEmptyRow(table);

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
                                        this.AddNewEmptyRow(table, false);

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
                if (exp.InnerException != null && exp.InnerException.InnerException != null)
                {
                    //if (exp.InnerException.InnerException.HResult == -2147024864)
                    //{
                    //    MessageBox.Show(this, "\"" + this.saveFileDialog1.FileName + "\" is already is use.\n\nPlease close it and generate report again.");
                    //}
                    //if (exp.InnerException.InnerException.HResult == -2147024891)
                    //{
                    //    MessageBox.Show(this, "You did not have rights to save file on selected location.\n\nPlease run as administrator.");
                    //}
                }                
                else
                {
                    MessageBox.Show(this, exp.Message);
                }
                
            }
            
        }

        private void AddMainHeading(Table table)
        {
            table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            Cell headingCell = new Cell(2, 4);
            headingCell.SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER);
            headingCell.SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3));
            headingCell.Add(new Paragraph("Attendance Report").SetFontSize(22F).SetBackgroundColor(new DeviceRgb(252, 213, 180)).SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 3)));
            table.AddCell(headingCell);
            
            table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
            table.AddCell(new Cell().Add(new Paragraph(string.Empty).SetFontSize(22F)).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 3)));
        }

        private void AddNewEmptyRow(Table table, bool removeBottomBorder = true)
        {
            table.StartNewRow();
            table.StartNewRow();

            if (removeBottomBorder)
            {
                table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
                table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            }
            else
            {
                table.AddCell(new Cell().
                    SetHeight(22F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(22F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(22F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(22F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(22F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(22F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(22F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(22F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));

                table.AddCell(new Cell().
                    SetHeight(22F).
                    SetBorderLeft(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderTop(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)).
                    SetBorderRight(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            }

        }

        private void AddDepartmentRow(Table table, string departmentName)
        {
            table.StartNewRow();
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
                    SetFontSize(11F)).
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
        }

        private void AddSectionRow(Table table, string sectionName)
        {
            table.StartNewRow();
            table.StartNewRow();

            table.AddCell(new Cell().SetHeight(22F).SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph("Section:").
                    SetFontSize(11F).
                    SetBold().
                    SetFontColor(new DeviceRgb(247, 150, 70))).
                SetHorizontalAlignment(iText.Layout.Properties.HorizontalAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE).
                SetHeight(22F).
                SetBorder(new iText.Layout.Borders.SolidBorder(iText.Kernel.Colors.Color.WHITE, 1)));
            table.AddCell(new Cell().
                    Add(new Paragraph(sectionName).
                    SetFontSize(11F)).
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
        }

        private void AddTableHeaderRow(Table table)
        {
            table.StartNewRow();
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
                    Add(new Paragraph("Crew").
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
                    Add(new Paragraph("CNIC Number").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph("Company Name").
                    SetFontSize(11F)).
                SetBackgroundColor(new DeviceRgb(253, 233, 217)).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.CENTER).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
        }

        private void AddTableDataRow(Table table, CardHolderInfo chl, bool altRow)
        {
            if (chl == null)
            {
                return;
            }

            

            table.StartNewRow();
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
                SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.OccurrenceTime.ToString()) ? string.Empty : chl.OccurrenceTime.ToString()).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.FirstName) ? string.Empty : chl.FirstName).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.PNumber) ? string.Empty : chl.PNumber).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.Crew) ? string.Empty : chl.Crew).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.Cadre) ? string.Empty : chl.Cadre).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.CNICNumber) ? string.Empty : chl.CNICNumber).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
            table.AddCell(new Cell().
                    Add(new Paragraph(string.IsNullOrEmpty(chl.Company) ? string.Empty : chl.Company).
                    SetFontSize(11F)).
                SetBackgroundColor(altRow ? new DeviceRgb(211, 211, 211) : iText.Kernel.Colors.Color.WHITE).
                SetBorder(new iText.Layout.Borders.SolidBorder(new DeviceRgb(247, 150, 70), 1)).
                SetTextAlignment(iText.Layout.Properties.TextAlignment.LEFT).
                SetVerticalAlignment(iText.Layout.Properties.VerticalAlignment.MIDDLE));
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
}
