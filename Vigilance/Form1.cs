using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Vigilance
{
    public partial class Form1 : Form
    {
        private static Object selectedItem;
        Microsoft.Office.Interop.Excel.Application theApp = Globals.ThisAddIn.Application;
        private static Microsoft.Office.Interop.Excel.Range selectedRange;

        public Form1()
        {
            InitializeComponent();
        }

        /// <summary>
        /// add condition to list
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            string rangeStr = textBox1.Text;
            double standard = 0;
            try
            {
                standard = Convert.ToDouble(textBox2.Text);
            }
            catch (Exception ex)
            {
                throw new Exception("输入错误:" + ex.Message);
            }

            string relationStr = comboBox1.SelectedItem.ToString();
            string conditionStr = rangeStr + " " + relationStr + " " + standard.ToString();

            if (rangeStr == null || textBox2 == null)
            {
                textBox1.Clear();
                textBox2.Clear();
                throw new Exception("信息不完整,请重新填写.");
            }

            //同时添加表和list
            ThisAddIn.Register(selectedRange, Relation.Equalto, standard);
            listBox1.Items.Add(conditionStr);
            NotifyChanges(selectedRange);

            textBox1.Clear();
            textBox2.Clear();
        }

        /// <summary>
        /// remove condition
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            Object selected = listBox1.SelectedItem;
            int indexSelect = listBox1.SelectedIndex;
            //MessageBox.Show(indexSelect.ToString());
            //MessageBox.Show(ThisAddIn.condList.Count.ToString());
            Condition cond = ThisAddIn.condList[indexSelect];

            if (selected == null)
            {
                throw new Exception("没有选中任何一个项.");
            }
            else
            {
                //同时删除表和list
                ThisAddIn.Unregister(cond);
                listBox1.Items.Remove(selected);
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            selectedItem = listBox1.SelectedItem;
        }

        /// <summary>
        /// select range
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button3_Click(object sender, EventArgs e)
        {
            selectedRange = theApp.Selection
                as Microsoft.Office.Interop.Excel.Range;
            textBox1.Text = selectedRange.AddressLocal;
        }

        private void NotifyChanges(Microsoft.Office.Interop.Excel.Range range)
        {
            Microsoft.Office.Interop.Excel.Worksheet nativeWorksheet =
                Globals.ThisAddIn.Application.ActiveSheet;
            
            nativeWorksheet.Change += new Microsoft.Office.Interop.Excel.DocEvents_ChangeEventHandler(changesRange_Change);
            if (nativeWorksheet != null)
            {
                Microsoft.Office.Tools.Excel.Worksheet vstoWorksheet =
                    Globals.Factory.GetVstoObject(nativeWorksheet);
            }
        }

        void changesRange_Change(Microsoft.Office.Interop.Excel.Range Target)
        {
            MessageBox.Show("Cell " + Target + " changed.");
        }
    }
}
