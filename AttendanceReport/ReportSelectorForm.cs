using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AttendanceReport
{
    public partial class ReportSelectorForm : Form
    {
        public ReportSelectorForm()
        {
            InitializeComponent();
        }

        private void btnLateArrivalReport_Click(object sender, EventArgs e)
        {
            Form1 form = new Form1(true);

            form.ShowDialog(this);
        }

        private void ReportSelectorForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (!LoginForm.mMainForm.Visible)
            {
                LoginForm.mMainForm.Close();
            }
        }

        private void btnAttendance_Click(object sender, EventArgs e)
        {
            Form1 form = new Form1(false);

            form.ShowDialog(this);
        }

        private void btnContractorWise_Click(object sender, EventArgs e)
        {
            ContractorWiseReportForm form = new ContractorWiseReportForm(false);

            form.ShowDialog(this);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ContractorWiseReportForm form = new ContractorWiseReportForm(true);

            form.ShowDialog(this);
        }

        private void btnCategoryWise_Click(object sender, EventArgs e)
        {
            CategoryWiseReportForm form = new CategoryWiseReportForm(false);

            form.ShowDialog(this);
        }

        private void btnCategoryWiseNotReturned_Click(object sender, EventArgs e)
        {
            CategoryWiseReportForm form = new CategoryWiseReportForm(true);

            form.ShowDialog(this);
        }

        private void btnBannedEntry_Click(object sender, EventArgs e)
        {
            BannedEntryReportForm form = new BannedEntryReportForm();

            form.ShowDialog(this);
        }

        private void btnAlarmedListReport_Click(object sender, EventArgs e)
        {
            AlarmedListReportForm form = new AlarmedListReportForm();

            form.ShowDialog(this);
        }

        private void btnAttendanceReport_Click(object sender, EventArgs e)
        {
            LMSAttendanceReportForm form = new LMSAttendanceReportForm();

            form.ShowDialog(this);
        }
    }
}
