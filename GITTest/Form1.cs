using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GITTest
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {
            //comment here
        }



        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btnGetDates_Click(object sender, EventArgs e)
        {
            List<string> Dates = new List<string>();
            //clear the list box
            listBoxDates.Items.Clear();
            //create the connection string
            string connectionString = @"Provider = Microsoft.ACE.OLEDB.12.0; Data Source = 'Databases\Data set 1.accdb'";
            using (OleDbConnection Connection = new OleDbConnection(connectionString))
            {

                Connection.Open();
                OleDbDataReader reader = null;
                OleDbCommand getDates = new OleDbCommand("SELECT [Order Date], [Ship Date] from Sheet1'", Connection);
                reader = getDates.ExecuteReader();
                while (reader.Read())
                {
                    Dates.Add(reader[0].ToString());
                    Dates.Add(reader[1].ToString());
                }
            }
            //create a new list for the formatted data
            List<string> DatesFormatted = new List<string>();
            
            foreach (string date in Dates)
            {
                //split the string on the whitespace and remove anything thats blank
                var dates = date.Split(new char[0], StringSplitOptions.RemoveEmptyEntries);
                //grab the first item(we know this is the date) and add it into our list
                DatesFormatted.Add(dates[0]);

            }
            
            //bind the list to the listbox
            listBoxDates.DataSource = DatesFormatted;
            //
            
        }
    }
}
