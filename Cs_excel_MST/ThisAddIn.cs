using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.IO;
using System.Windows.Forms;

namespace Cs_excel_MST
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //this.Application.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(WorkBookOpen);
            this.Application.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
            this.Application.WorkbookAfterSave += new Excel.AppEvents_WorkbookAfterSaveEventHandler(Application_WorkbookAfterSave);

            ((Excel.AppEvents_Event)this.Application).NewWorkbook += new Excel.AppEvents_NewWorkbookEventHandler(App_NewWorkbook);
            ((Excel.AppEvents_Event)this.Application).WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(App_WorkbookOpen);

            Form1 openForm = new Form1();
            openForm.TopMost = true;
            openForm.Show();


        }

        void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            Cancel = true; // Prevent the normal save process
            Wb.Application.EnableEvents = false;
            Wb.SaveAs(Wb.FullName, Password: "Timis-1864"); // Save with the original password
            Wb.Application.EnableEvents = true;
        }


        void Application_WorkbookAfterSave(Excel.Workbook Wb, bool Success)
        {
            string backupPath = "";
            try
            {
                Excel.Worksheet settingsSheet = Wb.Sheets["Settings"];
                Excel.Range backupPathCell = settingsSheet.Range["B1"];
                backupPath = backupPathCell.Value;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Could not fined value of path: {ex.Message}");

            }

            if (!Directory.Exists(backupPath))
            {
                MessageBox.Show("Folderul de backup nu a fost gasit. Contactati compartimentul informatica");
            }
            else
            {
                string backupFileName = $"{DateTime.Now.ToString("dd-MM")}-{Wb.Name}-backup";
                string backupFilePath = Path.Combine(backupPath, backupFileName);
                Wb.SaveCopyAs(backupFilePath);
            }
        }

        void App_WorkbookOpen(Excel.Workbook Wb)
        {
            // Check if this is a new workbook and take action
            // List of allowed workbook names (case-insensitive)
            string[] allowedWorkbooks = {
            "Registre_evidenta_asesizarilor_confirmarea_autorizarea_interceptarilor.xlsx",
            "Registre_evidenta_autorizatiilor_perchezitie_domiciliara_UP.xlsx"
            };

            // Check if the opened workbook is not in the list of allowed workbooks
            if (!allowedWorkbooks.Contains(Wb.Name, StringComparer.OrdinalIgnoreCase))
            {
                var result = System.Windows.Forms.MessageBox.Show(
                    $"Unauthorized workbook detected: {Wb.Name}. This workbook will now be closed.",
                    "Unauthorized Workbook",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning
                );

                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    // Close the workbook without saving changes
                    try
                    {
                        Wb.Close(SaveChanges: false);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show(
                            $"Error closing workbook: {ex.Message}",
                            "Error",
                            System.Windows.Forms.MessageBoxButtons.OK,
                            System.Windows.Forms.MessageBoxIcon.Error
                        );
                    }
                }
            }
        }

        void App_NewWorkbook(Excel.Workbook Wb)
        {
            MessageBox.Show("Creating new workbooks is disabled.");
            Wb.Close(false); // Close the new workbook without saving
        }


        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            this.Application.WorkbookBeforeSave -= new Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
            this.Application.WorkbookAfterSave -= new Excel.AppEvents_WorkbookAfterSaveEventHandler(Application_WorkbookAfterSave);
            ((Excel.AppEvents_Event)this.Application).NewWorkbook -= new Excel.AppEvents_NewWorkbookEventHandler(App_NewWorkbook);
            ((Excel.AppEvents_Event)this.Application).WorkbookOpen -= new Excel.AppEvents_WorkbookOpenEventHandler(App_WorkbookOpen);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
