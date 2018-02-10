using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Inventor;

namespace HCMToolsInventorAddIn
{
    public partial class InvBasicProgressForm : Form
    {
        Transaction oTrans;
        public InvBasicProgressForm(Inventor.Application app, Transaction CurrentTrans)
        {
            oTrans = CurrentTrans;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AddinGlobal.InventorApp.CommandManager.StopActiveCommand();
            this.Close();
        }
    }
}
