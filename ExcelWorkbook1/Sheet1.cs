using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelWorkbook1
{
    public partial class Sheet1
    {
        Microsoft.Office.Tools.Excel.NamedRange changesRange;

        private void Sheet1_Startup(object sender, System.EventArgs e)
        {
            SubscribeChanges();
        }

        private void Sheet1_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void SubscribeChanges()
        {
            changesRange = this.Controls.AddNamedRange(this.Range["B2", "E5"], "compositeRange");
            changesRange.Change += new Excel.DocEvents_ChangeEventHandler(changesRange_Change);

        }

        void changesRange_Change(Excel.Range Target)
        {
            Condition cond = new Condition();
            cond.relation = Relation.Greaterthan;
            cond.starndard = 30;
            
            string cellAddress = Target.get_Address(Excel.XlReferenceStyle.xlA1);
            if (Target.Value > 30)
            {
                System.Media.SystemSounds.Exclamation.Play();
                MessageBox.Show("Cell " + cellAddress + " changed.");
            }
            
        }

        #region VSTO Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Sheet1_Startup);
            this.Shutdown += new System.EventHandler(Sheet1_Shutdown);
        }

        #endregion

    }
}
