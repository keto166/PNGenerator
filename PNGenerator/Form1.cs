using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Threading;
using System.Collections;

namespace PNGenerator
{
    public partial class Form1 : Form
    {
        Excel.Application objApp;
        Excel.Workbooks objBooks;
        Excel.Sheets objSheets;
        Excel._Worksheet objSheet;
        Excel.Workbook objBook;
        Excel.Range objRange;

        Dictionary<String, int> cats;



        public Form1()
        {
            InitializeComponent();


        }

        private void Form1_Load_1(object sender, EventArgs e)
        {
            this.button1.Text = "GHeyyy";
            //MessageBox.Show("Heyy");
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {


            try
            {
                // Instantiate Excel and start a new workbook.
                ArrayList exInst = new ArrayList();  //List of existing Excel
                Process[] P_CESSES = Process.GetProcessesByName("EXCEL");
                for (int p_count = 0; p_count < P_CESSES.Length; p_count++)
                {
                    exInst.Add(P_CESSES[p_count].Id);
                        
                    

                }
                objApp = new Excel.Application();

                objBooks = objApp.Workbooks;
                objBook = objBooks.Open("Y:\\Mechanical\\Inventor\\iLogic_Plug_Ins\\iLogic_Drawings_Plug_Ins\\TestXL.xlsx");
                objSheets = objBook.Worksheets;
                objSheet = (Excel._Worksheet)objSheets.get_Item(1);




                objSheet.Cells[2, 1] = "Heyy";

                
                Boolean bEndFound = false;
                int n = 0;
                String tempString;
                int tempCount = 0;
                if (cats == null) { cats = new Dictionary<string, int>(); } else { cats.Clear(); }
                ComboBox _cbCats = cbCats;

                while (!bEndFound)
                {
                    objRange = ((Excel.Range)objSheet.Cells[3 + n, 4]);
                    tempString = objRange.Text;
                    if (tempString.Equals("")) {
                        bEndFound = true;
                    } else
                    {
                        objRange = ((Excel.Range)objSheet.Cells[3 + n, 5]);
                        tempCount = Convert.ToInt32(objRange.Value2);
                        cats.Add(tempString, tempCount);
                        n++;
                        _cbCats.Items.Add(tempString);
                    }
                }

                

                objBook.Save();
                objBook.Close(0);

                Boolean bFound = false;
                P_CESSES = Process.GetProcessesByName("EXCEL");
                for (int p_count = 0; p_count < P_CESSES.Length; p_count++)
                {
                    bFound = false;
                    foreach (int s in exInst)
                    {
                        if (s.Equals(P_CESSES[p_count].Id)) {
                            bFound = true;
                        }
                    }

                    if (!bFound)
                    {
                        P_CESSES[p_count].Kill();
                    }
                }

                while (System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objApp) != 0) { }
                while (System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objBooks) != 0) { }
                while (System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objBook) != 0) { }
                while (System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objSheets) != 0) { }
                while (System.Runtime.InteropServices.Marshal.FinalReleaseComObject(objSheet) != 0) { }
                objApp = null;
                objBooks = null;
                objBook = null;
                objSheets = null;
                objSheet = null;
                

                GC.Collect();
                GC.WaitForPendingFinalizers();

            }
            catch (Exception myError)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, myError.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, myError.Source);

                MessageBox.Show(errorMessage, "Error");
            }

            
        }

        private void cbCats_SelectedIndexChanged(object sender, EventArgs e)
        {
            String temp = (String)cbCats.SelectedItem;
            flCount.Text = cats[temp].ToString("D5");
        }
    }
}
