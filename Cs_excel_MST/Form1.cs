using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Cs_excel_MST
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.FormClosed += new FormClosedEventHandler(OpenWorkbookForm_FormClosed);
        }

        private void OpenWorkbook(string filePath, string password)
        {
            try
            {
                Globals.ThisAddIn.Application.Workbooks.Open(filePath, Password: password);
                // this.Hide(); // Close the form after opening the workbook

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to open workbook: {ex.Message}");
            }
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (!Globals.ThisAddIn.Application.Ready)
            {
                MessageBox.Show("Excel is currently busy. Please try again in a moment.");
                return;
            }
            OpenWorkbook(@"C:\Users\scaravelea\OneDrive\10.Career\03.DEPARTAMENT IT\registre2023\secrete\Registre_evidenta_asesizarilor_confirmarea_autorizarea_interceptarilor.xlsx", "Timis-1864");

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (!Globals.ThisAddIn.Application.Ready)
            {
                MessageBox.Show("Excel is currently busy. Please try again in a moment.");
                return;
            }
            OpenWorkbook(@"C:\Users\scaravelea\OneDrive\10.Career\03.DEPARTAMENT IT\registre2023\secrete\Registre_evidenta_autorizatiilor_perchezitie_domiciliara_UP.xlsx", "Timis-1864");
        }


        private void OpenWorkbookForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (Application.OpenForms.Count == 0)
            {
                Globals.ThisAddIn.Application.Quit();

            }
        }

    }
}
