using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text; 
using System.Windows.Forms;

namespace BillinSoft
{
    public partial class BillingSoft : Form
    {

        Microsoft.Office.Interop.Excel.Workbooks wrbks = null;
        Microsoft.Office.Interop.Excel.Workbook wrbk = null;
        Microsoft.Office.Interop.Excel.Worksheet wrst = null;

        Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
        private bool edit;

        public BillingSoft()
        {
            InitializeComponent();
           this.FormBorderStyle = FormBorderStyle.FixedSingle;

            this.MaximizeBox = false;

            excel.Application.Workbooks.Add(true);
            wrbks = excel.Workbooks;

            textBox2.Text = DateTime.Now.ToString("d/M/yyyy");
        }

        private void Form1_Load(object sender, EventArgs e)
        {



        }

        private void button1_Click(object sender, EventArgs e)
        {

            textBox1.TabStop = false;
            textBox2.TabStop = false;
            textBox3.TabStop = false;
            textBox4.TabStop = false;
            textBox5.TabStop = false;
            comboBox1.TabStop = false;



            string itemName = textBox6.Text;
            string qty = textBox7.Text;
            string amount = textBox8.Text;


            string[] row = { itemName, qty, amount };

            dataGridView1.Rows.Add(row);


            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";

            
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(!edit)      
            save();
            

            string st = textBox3.Text + textBox1.Text + ".xlsm";

            if(st== ".xlsm")
            {
                st = "temp.xlsm";
            }

           
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook wb = excelApp.Workbooks.Open(
            st,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // Get the first worksheet.
            // (Excel uses base 1 indexing, not base 0.)
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)wb.Worksheets[1];

            // Print out 1 copy to the default printer:

            bool userDidntCancel =
    excelApp.Dialogs[Microsoft.Office.Interop.Excel.XlBuiltInDialog.xlDialogPrint].Show(
        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
        Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            

            if (!userDidntCancel)
            {
                excelApp.Quit();
                return;
            }

            excelApp.Quit();

        }

        private void save()
        {
            try
            {
                wrbk = wrbks.Open("C:\\Program Files (x86)\\WC\\Setup\\bill.xlsm");
                wrst = (Microsoft.Office.Interop.Excel.Worksheet)wrbk.Worksheets[1];
            }

            catch (Exception ex)
            {
                MessageBox.Show("can't load sample bill!!!");
                return;
            }

            for (int rows = 0; rows < dataGridView1.Rows.Count - 1; rows++)
            {
                string str = 'A' + (rows + 11).ToString();
                wrst.Range[str].Value = (rows+1).ToString();

                str = 'B' + (rows + 11).ToString();
                wrst.Range[str].Value = dataGridView1.Rows[rows].Cells[0].Value.ToString();

                str = 'D' + (rows + 11).ToString();
                wrst.Range[str].Value = dataGridView1.Rows[rows].Cells[1].Value.ToString();

                str = 'E' + (rows + 11).ToString();
                wrst.Range[str].Value = dataGridView1.Rows[rows].Cells[2].Value.ToString();

            }

            string str1 = "A6";
             wrst.Range[str1].Value = textBox1.Text;

            str1 = "B6";
            wrst.Range[str1].Value = textBox2.Text;

            str1 = "D6";
            wrst.Range[str1].Value = textBox3.Text;

            str1 = "E6";

            string adr = textBox4.Text;
            adr = adr.Replace(System.Environment.NewLine," ");
            Console.Write(adr);
            wrst.Range[str1].Value = adr;

            str1 = "F6";
            wrst.Range[str1].Value = textBox5.Text;

            str1 = "A9";
            wrst.Range[str1].Value = comboBox1.Text;

            string st =  textBox3.Text + textBox1.Text + ".xlsm";

            if (st == ".xlsm")
            {
                st = "temp.xlsm";
            }

            excel.ActiveWorkbook.SaveCopyAs(st);
            excel.ActiveWorkbook.Saved = true;


            GC.Collect();
            GC.WaitForPendingFinalizers();

            excel.Quit();

            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            save();
            MessageBox.Show("File is Saved in Documents folder");

        }

        private void button4_Click(object sender, EventArgs e)
        {
            save();

            string st = textBox3.Text + textBox1.Text + ".xlsm";

            if (st == ".xlsm")
            {
                st = "temp.xlsm";
            }

            edit = true;

            button3.Hide();
            Process.Start(Path.Combine(Environment.ExpandEnvironmentVariables("%userprofile%"), "Documents") + "\\" + st);
        }
    }
}
