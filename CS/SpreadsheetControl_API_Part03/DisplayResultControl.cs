using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraSpreadsheet;

namespace SpreadsheetControl_API_Part03
{
    public partial class DisplayResultControl : UserControl
    {
        public DisplayResultControl()
        {
            InitializeComponent();
        }

        public SpreadsheetControl Speadsheet { get { return spreadsheetControl1; } }

}
    }
