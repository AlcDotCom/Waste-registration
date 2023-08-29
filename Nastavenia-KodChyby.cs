using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Waste_registration.Properties;

namespace Waste_registration
{
    public partial class KodChyby : Form
    {
        public KodChyby()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Nastavenia nextForm = new Nastavenia();
            Hide();
            nextForm.ShowDialog();
            Close();
        }
        private void KodChyby_Load(object sender, EventArgs e)
        {
         PopulateDataGridView();
        }
        void PopulateDataGridView()
        {
            using (SqlConnection sqlCon = new SqlConnection(Connection.ConnectionString))
            {
                try { 
                sqlCon.Open();
                SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT * FROM SM30detail", sqlCon);
                DataTable dtbl = new DataTable();
                sqlDa.Fill(dtbl);
                dataGridView1.DataSource = dtbl;

                dtbl.DefaultView.RowFilter = string.Format("KodChyby <> '{0}'", "");

                dataGridView1.Columns["AdresaUlozeniaCSV"].Visible = false;
                dataGridView1.Columns["Rezerva1"].Visible = false;
                dataGridView1.Columns["Rezerva2"].Visible = false;
                dataGridView1.Columns["Rezerva3"].Visible = false;
                dataGridView1.Columns["MenoOdpisovacaOdpadu"].Visible = false;
                } catch (Exception) {}
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                using (SqlConnection sqlCon = new SqlConnection(Connection.ConnectionString))
                {
                    sqlCon.Open();
                    DataGridViewRow dgvRow = dataGridView1.CurrentRow;
                    SqlCommand sqlCmd = new SqlCommand("SM30detailEDITcode", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Parameters.AddWithValue("@ID", dgvRow.Cells["txtID"].Value == DBNull.Value ? "" : dgvRow.Cells["txtID"].Value.ToString());
                    sqlCmd.Parameters.AddWithValue("@KodChyby", dgvRow.Cells["txtKodChyby"].Value == DBNull.Value ? "" : dgvRow.Cells["txtKodChyby"].Value.ToString());
                    sqlCmd.Parameters.AddWithValue("@PopisChyby", dgvRow.Cells["txtPopisChyby"].Value == DBNull.Value ? "" : dgvRow.Cells["txtPopisChyby"].Value.ToString());
                    sqlCmd.Parameters.AddWithValue("@Oddelenie", dgvRow.Cells["txtOddelenie"].Value == DBNull.Value ? "" : dgvRow.Cells["txtOddelenie"].Value.ToString());
                    sqlCmd.Parameters.AddWithValue("@DatumUpravyPridania", DateTime.Now);
                    sqlCmd.Parameters.AddWithValue("@UpravilPridal", Environment.MachineName);
                    sqlCmd.ExecuteNonQuery();
                    PopulateDataGridView();
                }
            }
        }
    }
}
