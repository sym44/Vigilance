using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace Vigilance
{
    public partial class ThisAddIn
    {
        public static List<Condition> condList = null; // remember the conditions

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            condList = new List<Condition>();
        }

        public static void Register(Excel.Range range, Relation relation, double standard)
        {
            Condition cond = new Condition();
            cond.range = range;
            cond.relation = relation;
            condList.Add(cond);
        }

        public static void Unregister(Condition cond)
        {
            condList.Remove(cond);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
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
