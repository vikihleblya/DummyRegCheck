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
using OfficeOpenXml;

namespace RegigCheck
{
    public partial class Form1 : Form
    {
        ExcelPackage excel = new ExcelPackage();
        String fullName = "";
        String creator = "";
        String service = "";
        String date = "";

        public Form1() {
            InitializeComponent();
        }

        private void Create_Click(object sender, EventArgs e) {
            if (checkTextfields()) {
                createExcelFile();
            } else {
                label5.Visible = true;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e) {
            this.fullName = textBox1.Text;
        }

        private void textBox2_TextChanged(object sender, EventArgs e) {
            this.creator = textBox2.Text;
        }

        private void textBox3_TextChanged(object sender, EventArgs e) {
            this.service = textBox3.Text;
        }

        private void textBox4_TextChanged(object sender, EventArgs e) {
            this.date = textBox4.Text;
        }


        private void createExcelFile() {
            excel.Workbook.Worksheets.Add("Worksheet1");

            var headerRow = new List<string[]>() {
                new string[] { fullName, creator, service, date }
            };

            string headerRange = "A1:" + Char.ConvertFromUtf32(headerRow[0].Length + 64) + "1";

            var worksheet = excel.Workbook.Worksheets["Worksheet1"];

            worksheet.Cells[headerRange].LoadFromArrays(headerRow);

            FileInfo excelFile = new FileInfo($@"C:\Users\Виктор\Desktop\{fullName}{date}.xlsx");
            excel.SaveAs(excelFile);
        }

        private bool checkTextfields() {
            if (fullName.Length == 0 & service.Length == 0 & creator.Length == 0 & date.Length == 0) {
                return false;
            }
            return true;
        }
    }
}
