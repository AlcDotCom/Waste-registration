using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Waste_registration.Properties;

namespace Waste_registration
{
    public partial class Adresa : Form
    {
        public Adresa()
        {
            InitializeComponent();
            Settings.Default.AdresaOverenie = "";
            Settings.Default.Save();
            if (Settings.Default.StatusAdresa == "1")
            { checkBox1.Checked = true;
                dataGridView1.Visible = false;
                label1.Visible = false;
            }
            OverenieZadanejAdresy();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Dispose();
            Nastavenia nextForm = new Nastavenia();
            Hide();
            nextForm.ShowDialog();
            Close();
        }
        public new void Dispose()
        {
            pictureBox1.Image.Dispose();
            pictureBox1.Image = null;
        }

        private void Adresa_Load(object sender, EventArgs e)
        {
            PopulateDataGridView();
        }
        void PopulateDataGridView()
        {
            using (SqlConnection sqlCon = new SqlConnection(Connection.ConnectionString))
            {
                try
                {
                    sqlCon.Open();
                    SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT * FROM SM30detail", sqlCon);
                    DataTable dtbl = new DataTable();
                    sqlDa.Fill(dtbl);
                    dataGridView1.DataSource = dtbl;
                    dtbl.DefaultView.RowFilter = string.Format("AdresaUlozeniaCSV <> '{0}'", "");
                    dataGridView1.Columns["KodChyby"].Visible = false;
                    dataGridView1.Columns["PopisChyby"].Visible = false;
                    dataGridView1.Columns["Rezerva1"].Visible = false;
                    dataGridView1.Columns["Rezerva2"].Visible = false;
                    dataGridView1.Columns["Rezerva3"].Visible = false;
                    dataGridView1.Columns["Oddelenie"].Visible = false;
                    dataGridView1.Columns["MenoOdpisovacaOdpadu"].Visible = false; 
                }
                catch (Exception) { }
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
                    SqlCommand sqlCmd = new SqlCommand("SM30detailEDITaddress", sqlCon);
                    sqlCmd.CommandType = CommandType.StoredProcedure;
                    sqlCmd.Parameters.AddWithValue("@ID", dgvRow.Cells["txtID"].Value == DBNull.Value ? "" : dgvRow.Cells["txtID"].Value.ToString());
                    sqlCmd.Parameters.AddWithValue("@AdresaUlozeniaCSV", dgvRow.Cells["txtAdresaUlozeniaCSV"].Value == DBNull.Value ? "" : dgvRow.Cells["txtAdresaUlozeniaCSV"].Value.ToString());
                    sqlCmd.Parameters.AddWithValue("@DatumUpravyPridania", DateTime.Now);
                    sqlCmd.Parameters.AddWithValue("@UpravilPridal", Environment.MachineName);
                    sqlCmd.ExecuteNonQuery();
                    PopulateDataGridView();
                    OverenieZadanejAdresy();
                }
            }
        }
        void OverenieZadanejAdresy()
        {
            try
            {
                if (Settings.Default.StatusAdresa == "0")
                {
                    using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                    {
                        SqlCommand sqlCmd = new SqlCommand("SELECT AdresaUlozeniaCSV FROM SM30detail WHERE AdresaUlozeniaCSV <> ''", sqlConnection);
                        sqlConnection.Open();
                        SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                        while (sqlReader.Read())
                        {
                            Settings.Default.AdresaOverenie = sqlReader["AdresaUlozeniaCSV"].ToString();
                            Settings.Default.Save();
                        }
                        sqlReader.Close();
                    }
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Adresu uloženia nebolo možné načítať, overte pripojenie so serverom");
            }
            if (Settings.Default.AdresaOverenie != "")
            {
                dataGridView1.AllowUserToAddRows = false;
            }
            else
            {
                dataGridView1.AllowUserToAddRows = true;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                Process.Start(Settings.Default.AdresaOverenie);
            }
            catch (Win32Exception win32Exception)
            {
                MessageBox.Show(win32Exception.Message);
            }
        }

        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                Settings.Default.AdresaOverenie = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                Settings.Default.StatusAdresa = "1";
                Settings.Default.Save();
                dataGridView1.Visible = false;
                label1.Visible = false;
            }
            else
            {
                Settings.Default.StatusAdresa = "0";
                Settings.Default.Save();
                dataGridView1.Visible = true;
                label1.Visible = true;
            }
        }
    }
}
