using AttendanceReport.CCFTCentral;
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
                        
            this.mDepartmentTask.Start();
            this.mSectionTask.Start();
            this.mCrewTask.Start();

            //List<string> lstDepartments = departmentTask.Result;
            //this.cbxDepartments.Items.Add(string.Empty);
            //this.cbxDepartments.Items.AddRange(lstDepartments.ToArray());

            //List<string> lstSections = sectionTask.Result;
            //this.cbxSections.Items.Add(string.Empty);
            //this.cbxSections.Items.AddRange(lstSections.ToArray());

            //List<string> lstCrews = crewTask.Result;
            //this.cbxCrew.Items.Add(string.Empty);
            //this.cbxCrew.Items.AddRange(lstCrews.ToArray());
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Cursor currentCursor = Cursor.Current;
            try
            {
                
                Cursor.Current = Cursors.WaitCursor;
                this.mData = null;

                DateTime fromDate = this.dtpFromDate.Value;
                DateTime toDate = this.dtpToDate.Value;

                toDate = toDate.AddDays(1);
                
                
                       

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
                                    if (!string.IsNullOrEmpty(filterByCardNumber) )
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
                                            Crew = crew
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
                                            this.mData.Add(department,dict);
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



                    using (ExcelPackage ex = new ExcelPackage())
                    {
                        ExcelWorksheet work = ex.Workbook.Worksheets.Add("Attendence Report");

                        work.View.ShowGridLines = false;

                        work.Column(2).Width = 18.14;
                        work.Column(3).Width = 25.29;
                        work.Column(4).Width = 54;
                        work.Column(5).Width = 15.14;

                        //Heading
                        work.Cells["D2:D3"].Merge = true;
                        work.Cells["D2:D3"].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        work.Cells["D2:D3"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(252, 213, 180));
                        work.Cells["D2:D3"].Style.Font.Size = 22;
                        work.Cells["D2:D3"].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                        work.Cells["D2:D3"].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;
                        work.Cells["D2:D3"].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        work.Cells["D2:D3"].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        work.Cells["D2:D3"].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        work.Cells["D2:D3"].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                        work.Cells["D2:D3"].Style.Border.Top.Color.SetColor(Color.FromArgb(247, 150, 70));
                        work.Cells["D2:D3"].Style.Border.Bottom.Color.SetColor(Color.FromArgb(247, 150, 70));
                        work.Cells["D2:D3"].Style.Border.Left.Color.SetColor(Color.FromArgb(247, 150, 70));
                        work.Cells["D2:D3"].Style.Border.Right.Color.SetColor(Color.FromArgb(247, 150, 70));
                        work.Cells["D2:D3"].Value = "Attendance Report";

                        //Sections and Data

                        int row = 6;

                        foreach (KeyValuePair<string, Dictionary<string, List<CardHolderInfo>>> department in this.mData)
                        {
                            if (department.Value == null)
                            {
                                continue;
                            }

                            //Department
                            work.Cells[row, 2].Style.Font.Bold = true;
                            work.Cells[row, 2].Style.Font.Color.SetColor(Color.FromArgb(247, 150, 70));
                            work.Cells[row, 2, row, 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                            work.Cells[row, 2, row, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                            work.Cells[row, 2].Value = "Department:";
                            work.Cells[row, 3].Value = department.Key;
                            work.Row(row).Height = 20;

                            row++;

                            foreach (KeyValuePair<string, List<CardHolderInfo>> section in department.Value)
                            {
                                //Section
                                work.Cells[row, 2].Style.Font.Bold = true;
                                work.Cells[row, 2].Style.Font.Color.SetColor(Color.FromArgb(247, 150, 70));
                                work.Cells[row, 2, row, 3].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                work.Cells[row, 2, row, 3].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                work.Cells[row, 2].Value = "Section:";
                                work.Cells[row, 3].Value = section.Key;
                                work.Row(row).Height = 20;

                                //Data
                                row++;
                                row++;

                                work.Cells[row, 2, row, 6].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                work.Cells[row, 2, row, 6].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                work.Cells[row, 2, row, 6].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                work.Cells[row, 2, row, 6].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                            
                                work.Cells[row, 2, row, 6].Style.Border.Top.Color.SetColor(Color.FromArgb(247, 150, 70));
                                work.Cells[row, 2, row, 6].Style.Border.Bottom.Color.SetColor(Color.FromArgb(247, 150, 70));
                                work.Cells[row, 2, row, 6].Style.Border.Left.Color.SetColor(Color.FromArgb(247, 150, 70));
                                work.Cells[row, 2, row, 6].Style.Border.Right.Color.SetColor(Color.FromArgb(247, 150, 70));

                                work.Cells[row, 2, row, 6].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                work.Cells[row, 2, row, 6].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(253, 233, 217));
                                work.Cells[row, 2, row, 6].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                                work.Cells[row, 2, row, 6].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                                work.Cells[row, 2].Value = "Card Number";
                                work.Cells[row, 3].Value = "Occurrance Time";
                                work.Cells[row, 4].Value = "First Name";
                                work.Cells[row, 5].Value = "P-Number";
                                work.Cells[row, 6].Value = "Crew";
                                work.Row(row).Height = 20;

                                for (int i = 0; i < section.Value.Count; i++)
                                {
                                    row++;
                                    work.Cells[row, 2, row, 6].Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    work.Cells[row, 2, row, 6].Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    work.Cells[row, 2, row, 6].Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                                    work.Cells[row, 2, row, 6].Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;

                                    work.Cells[row, 2, row, 6].Style.Border.Top.Color.SetColor(Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 2, row, 6].Style.Border.Bottom.Color.SetColor(Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 2, row, 6].Style.Border.Left.Color.SetColor(Color.FromArgb(247, 150, 70));
                                    work.Cells[row, 2, row, 6].Style.Border.Right.Color.SetColor(Color.FromArgb(247, 150, 70));

                                    if (i % 2 == 0)
                                    {
                                        work.Cells[row, 2, row, 6].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                        work.Cells[row, 2, row, 6].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                                    }

                                    CardHolderInfo chl = section.Value[i];
                                    work.Cells[row, 2].Value = chl.CardNumber;
                                    work.Cells[row, 3].Value = chl.OccurrenceTime.ToString();
                                    work.Cells[row, 4].Value = chl.FirstName;
                                    work.Cells[row, 5].Value = chl.PNumber;
                                    work.Cells[row, 6].Value = chl.Crew;

                                    //if (chl.PNumber == "Late")
                                    //{
                                    //    work.Cells[row, 5].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    //    work.Cells[row, 5].Style.Fill.BackgroundColor.SetColor(Color.Red);
                                    //}

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

    }
}
