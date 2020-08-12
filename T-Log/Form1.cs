using OfficeOpenXml;
using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace T_Log
{
    public partial class Form1 : Form
    {
        //global variables of this class
        private string _fileName = "";

        private bool _caseExist = true;
        private int _lastRow = 1;
        private readonly Stopwatch _stopWatch;

        public Form1()
        {
            InitializeComponent();
            _stopWatch = new Stopwatch();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            AutoFillDate();
            PanelReview(true, false, false);
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            var openFile = new OpenFileDialog();

            if (openFile.ShowDialog() == DialogResult.OK)
            {
                _fileName = openFile.FileName;
                tbFile.Text = _fileName;
            }

            PanelReview(true, false, true);
            ControlReview(true, false, false);
            LoadDropdownValues();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            PanelReview(false, true, true);
            ControlReview(false, true, false);
            _stopWatch.Start();
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            PanelReview(false, true, true);
            ControlReview(false, false, true);

            _stopWatch.Stop();
            var duration = Math.Round(Convert.ToDecimal(_stopWatch.Elapsed.TotalHours), 2);
            tbDuration.Text = duration.ToString();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            WriteToExcelFile();
            ResetTask();
            PanelReview(true, false, true);
            ControlReview(true, false, false);
        }

        private void cbNA_CheckedChanged(object sender, EventArgs e)
        {
            if (cbNA.Checked)
            {
                tbCaseNo.Enabled = false;
                _caseExist = false;
            }
            else
            {
                tbCaseNo.Enabled = true;
                _caseExist = true;
            }
        }

        private void WriteToExcelFile()
        {
            var file = new FileInfo(_fileName);

            using var excel = new ExcelPackage(file);
            var wks = excel.Workbook.Worksheets[0];
            _lastRow = wks.Dimension.End.Row;

            wks.Cells[_lastRow + 1, 1].Value = _caseExist ? tbCaseNo.Text : "n/a";

            wks.Cells[_lastRow + 1, 2].Value = cbCustomer.SelectedItem.ToString();
            wks.Cells[_lastRow + 1, 3].Value = cbArea.SelectedItem.ToString();
            wks.Cells[_lastRow + 1, 4].Value = tbDescription.Text;
            wks.Cells[_lastRow + 1, 5].Value = tbDate.Text;
            wks.Cells[_lastRow + 1, 6].Value = tbDuration.Text;

            excel.Save();
        }

        private void ResetTask()
        {
            tbCaseNo.Enabled = true;
            cbNA.Enabled = true;
            cbNA.Checked = false;

            tbCaseNo.Text = "";
            cbArea.SelectedIndex = -1;
            cbCustomer.SelectedIndex = -1;
            tbDescription.Text = "";
            tbDuration.Text = "";
        }

        private void PanelReview(bool panFileEnable, bool panTaskEnable, bool panControlsEnable)
        {
            panFile.Enabled = panFileEnable;
            panTask.Enabled = panTaskEnable;
            panControls.Enabled = panControlsEnable;
        }

        private void ControlReview(bool startEnable, bool stopEnable, bool saveEnable)
        {
            btnStart.Enabled = startEnable;
            btnStop.Enabled = stopEnable;
            btnSave.Enabled = saveEnable;
        }

        private void AutoFillDate() => tbDate.Text = DateTime.Today.ToString("dd-MMM-yy");

        private void LoadDropdownValues()
        {
            var file = new FileInfo(_fileName);

            using var excel = new ExcelPackage(file);
            var wks = excel.Workbook.Worksheets[1];
            _lastRow = wks.Dimension.End.Row;

            for (var i = 2; i < _lastRow + 1; i++)
            {
                //customer will always be in Sheet2/column2
                var customerCellValue = wks.Cells[i, 2].Value;
                //area will always be in Sheet2/column2
                var areaCellValue = wks.Cells[i, 3].Value;

                if (customerCellValue != null)
                {
                    cbCustomer.Items.Add(customerCellValue);
                }

                if (areaCellValue != null)
                {
                    cbArea.Items.Add(areaCellValue);
                }
            }
        }
    }
}