using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using System.IO;

namespace T_Log
{
    public partial class Form1 : Form
    {
        //global variables of this class
        string fileName = "";
        bool taskStarted = false;
        bool caseExist = true;
        int lastRow = 1;

        Stopwatch stopWatch = new Stopwatch();


        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            AutoFillDate();
            PanelReview(true, false, false);
        }



        private void btnOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();

            if (openFile.ShowDialog() == DialogResult.OK)
            {
                fileName = openFile.FileName;
                tbFile.Text = fileName;
            }

            PanelReview(true, false, true);
            ControlReview(true, false, false);
            LoadDropdownValues();
        }

       
        private void btnStart_Click(object sender, EventArgs e)
        {
            PanelReview(false, true, true);
            ControlReview(false, true, false);
            taskStarted = true;
            stopWatch.Start();
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            PanelReview(false, true, true);
            ControlReview(false, false, true);
            taskStarted = false;

            stopWatch.Stop();
            decimal _duration = Math.Round(Convert.ToDecimal(stopWatch.Elapsed.TotalHours), 2);
            tbDuration.Text = _duration.ToString();


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
            if (cbNA.Checked == true)
            { 
                tbCaseNo.Enabled = false;
                caseExist = false;
            }
            else
            { 
                tbCaseNo.Enabled = true;
                caseExist = true;
            }
        }

        private void WriteToExcelFile()
        {
            FileInfo file = new FileInfo(fileName);

            using (ExcelPackage excel = new ExcelPackage(file))
            {

                ExcelWorksheet wks = excel.Workbook.Worksheets[0];
                lastRow = wks.Dimension.End.Row;


                if (caseExist == true)
                { wks.Cells[lastRow + 1, 1].Value = tbCaseNo.Text; }
                else
                { wks.Cells[lastRow + 1, 1].Value = "n/a"; }

                wks.Cells[lastRow + 1, 2].Value = cbCustomer.SelectedItem.ToString();
                wks.Cells[lastRow + 1, 3].Value = cbArea.SelectedItem.ToString();
                wks.Cells[lastRow + 1, 4].Value = tbDescription.Text;
                wks.Cells[lastRow + 1, 5].Value = tbDate.Text;
                wks.Cells[lastRow + 1, 6].Value = tbDuration.Text;

                excel.Save();
            }
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

        private void AutoFillDate()
        {
            tbDate.Text = DateTime.Today.ToString("dd-MMM-yy");
        }

        private void LoadDropdownValues()
        {
            FileInfo file = new FileInfo(fileName);

            using (ExcelPackage excel = new ExcelPackage(file))
            {

                ExcelWorksheet wks = excel.Workbook.Worksheets[1];
                lastRow = wks.Dimension.End.Row;

                for (int i = 2; i < lastRow + 1; i++)
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
}
