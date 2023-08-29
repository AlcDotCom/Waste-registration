using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Waste_registration.Properties;
using Excel = Microsoft.Office.Interop.Excel;

namespace Waste_registration
{
    public partial class Nastavenia__UpravaDat : Form
    {
        public Nastavenia__UpravaDat()
        {
            InitializeComponent();
            {
                try  //počiatočný ping serveru
                {
                    using (SqlConnection connection = new SqlConnection(Connection.ConnectionString))
                        connection.Open();
                }
                catch (Exception)
                {
                    MessageBox.Show("Spojenie so serverom zlyhalo, dáta nemožno zobraziť");
                    return;
                }
            }

            comboBox1.Items.Add("-Všetky-");
            try   //načíta oddelenia
            {
                using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                {
                    SqlCommand sqlCmd = new SqlCommand("SELECT Oddelenie FROM SM30 WHERE Aktiviny1Vymazany0 = '1'", sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlReader = sqlCmd.ExecuteReader();

                    while (sqlReader.Read())
                    {
                        comboBox1.Items.Add(sqlReader["Oddelenie"].ToString());
                    }
                    sqlReader.Close();
                }
            }
            catch (Exception)
            { }
            object[] distinctItems1 = (from object o in comboBox1.Items select o).Distinct().ToArray();
            comboBox1.Items.Clear();
            comboBox1.Items.AddRange(distinctItems1);
            comboBox1.SelectedIndex = 0;

            comboBox2.Items.Add("-Všetky-");
            try   //načíta projekty
            {
                using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                {
                    SqlCommand sqlCmd = new SqlCommand("SELECT Projekt FROM SM30 WHERE Aktiviny1Vymazany0 = '1'", sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlReader = sqlCmd.ExecuteReader();

                    while (sqlReader.Read())
                    {
                        comboBox2.Items.Add(sqlReader["Projekt"].ToString());
                    }
                    sqlReader.Close();
                }
            }
            catch (Exception)
            { }
            object[] distinctItems2 = (from object o in comboBox2.Items select o).Distinct().ToArray();
            comboBox2.Items.Clear();
            comboBox2.Items.AddRange(distinctItems2);
            comboBox2.SelectedIndex = 0;

            comboBox3.Items.Add("-Všetky-");
            try   //načíta výrobky
            {
                using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                {
                    SqlCommand sqlCmd = new SqlCommand("SELECT Vyrobok FROM SM30 WHERE Aktiviny1Vymazany0 = '1'", sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlReader = sqlCmd.ExecuteReader();

                    while (sqlReader.Read())
                    {
                        comboBox3.Items.Add(sqlReader["Vyrobok"].ToString());
                    }
                    sqlReader.Close();
                }
            }
            catch (Exception)
            { }
            object[] distinctItems3 = (from object o in comboBox3.Items select o).Distinct().ToArray();
            comboBox3.Items.Clear();
            comboBox3.Items.AddRange(distinctItems3);
            comboBox3.SelectedIndex = 0;
        }

        private void cbxDesign_DrawItem(object sender, DrawItemEventArgs e)  //zarovnanie dát v cbx na stred
        {
            ComboBox cbx = sender as ComboBox;
            if (cbx != null)
            {
                e.DrawBackground();
                if (e.Index >= 0)
                {
                    StringFormat sf = new StringFormat();
                    sf.LineAlignment = StringAlignment.Center;
                    sf.Alignment = StringAlignment.Center;
                    Brush brush = new SolidBrush(cbx.ForeColor);
                    if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                        brush = SystemBrushes.HighlightText;
                    e.Graphics.DrawString(cbx.Items[e.Index].ToString(), cbx.Font, brush, e.Bounds, sf);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Nastavenia nextForm = new Nastavenia();
            Hide();
            nextForm.ShowDialog();
            Close();
        }

        private void Nastavenia__UpravaDat_Load(object sender, EventArgs e)
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
                    SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT * FROM SM30", sqlCon);
                    DataTable dtbl = new DataTable();
                    sqlDa.Fill(dtbl);
                    dataGridView1.DataSource = dtbl;

                    if (checkBox1.Checked)
                    {
                        dtbl.DefaultView.RowFilter = string.Format("Aktiviny1Vymazany0 LIKE '{0}'", "0");
                        dataGridView1.Columns["txtVymazal"].Visible = true;
                        dataGridView1.Columns["txtDatumVymazu"].Visible = true;
                        dataGridView1.Sort(dataGridView1.Columns[16], ListSortDirection.Ascending);
                        comboBox1.SelectedIndex = 0;
                        comboBox2.SelectedIndex = 0;
                        comboBox3.SelectedIndex = 0;
                        comboBox1.Visible = false;
                        comboBox2.Visible = false;
                        comboBox3.Visible = false;
                        dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;
                        return;
                    }
                    else
                    {
                        dtbl.DefaultView.RowFilter = string.Format("Aktiviny1Vymazany0 LIKE '{0}'", "1");
                        dataGridView1.Columns["txtVymazal"].Visible = false;
                        dataGridView1.Columns["txtDatumVymazu"].Visible = false;
                        comboBox1.Visible = true;
                        comboBox2.Visible = true;
                        comboBox3.Visible = true;
                    }

                    if (comboBox1.SelectedIndex > 0 && comboBox2.SelectedIndex == 0 && comboBox3.SelectedIndex == 0)  //oddelenie
                    {
                        dtbl.DefaultView.RowFilter = string.Format("Oddelenie = '{0}' AND Aktiviny1Vymazany0 =  '{1}'", comboBox1.SelectedItem.ToString(),"1");

                        comboBox2.Items.Clear();  //aktualizuje projekty
                        comboBox2.Items.Add("-Všetky-");
                        try
                        {
                            using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                            {
                                SqlCommand sqlCmd = new SqlCommand("SELECT Projekt FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Oddelenie ='" + comboBox1.SelectedItem + "'", sqlConnection);
                                sqlConnection.Open();
                                SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                                while (sqlReader.Read())
                                {
                                    comboBox2.Items.Add(sqlReader["Projekt"].ToString());
                                }
                                sqlReader.Close();
                            }
                        }
                        catch (Exception)
                        { }
                        object[] distinctItems2 = (from object o in comboBox2.Items select o).Distinct().ToArray();
                        comboBox2.Items.Clear();
                        comboBox2.Items.AddRange(distinctItems2);
                        comboBox2.SelectedIndex = 0;

                        comboBox3.Items.Clear();  //aktualizuje vyrobky
                        comboBox3.Items.Add("-Všetky-");
                        try
                        {
                            using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                            {
                                SqlCommand sqlCmd = new SqlCommand("SELECT Vyrobok FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Oddelenie ='" + comboBox1.SelectedItem + "'", sqlConnection);
                                sqlConnection.Open();
                                SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                                while (sqlReader.Read())
                                {
                                    comboBox3.Items.Add(sqlReader["Vyrobok"].ToString());
                                }
                                sqlReader.Close();
                            }
                        }
                        catch (Exception)
                        { }
                        object[] distinctItems3 = (from object o in comboBox3.Items select o).Distinct().ToArray();
                        comboBox3.Items.Clear();
                        comboBox3.Items.AddRange(distinctItems3);
                        comboBox3.SelectedIndex = 0;
                    }

                    if (comboBox1.SelectedIndex == 0 && comboBox2.SelectedIndex > 0 && comboBox3.SelectedIndex == 0)  //projekt
                    {
                        dtbl.DefaultView.RowFilter = string.Format("Projekt = '{0}' AND Aktiviny1Vymazany0 =  '{1}'", comboBox2.SelectedItem.ToString(),"1");

                        comboBox1.Items.Clear();  //aktualizuje oddelenia
                        comboBox1.Items.Add("-Všetky-");
                        try
                        {
                            using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                            {
                                SqlCommand sqlCmd = new SqlCommand("SELECT Oddelenie FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Projekt ='" + comboBox2.SelectedItem + "'", sqlConnection);
                                sqlConnection.Open();
                                SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                                while (sqlReader.Read())
                                {
                                    comboBox1.Items.Add(sqlReader["Oddelenie"].ToString());
                                }
                                sqlReader.Close();
                            }
                        }
                        catch (Exception)
                        { }
                        object[] distinctItems1 = (from object o in comboBox1.Items select o).Distinct().ToArray();
                        comboBox1.Items.Clear();
                        comboBox1.Items.AddRange(distinctItems1);
                        comboBox1.SelectedIndex = 0;

                        comboBox3.Items.Clear();  //aktualizuje vyrobky
                        comboBox3.Items.Add("-Všetky-");
                        try
                        {
                            using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                            {
                                SqlCommand sqlCmd = new SqlCommand("SELECT Vyrobok FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Projekt ='" + comboBox2.SelectedItem + "'", sqlConnection);
                                sqlConnection.Open();
                                SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                                while (sqlReader.Read())
                                {
                                    comboBox3.Items.Add(sqlReader["Vyrobok"].ToString());
                                }
                                sqlReader.Close();
                            }
                        }
                        catch (Exception)
                        { }
                        object[] distinctItems3 = (from object o in comboBox3.Items select o).Distinct().ToArray();
                        comboBox3.Items.Clear();
                        comboBox3.Items.AddRange(distinctItems3);
                        comboBox3.SelectedIndex = 0;
                    }

                    if (comboBox1.SelectedIndex == 0 && comboBox2.SelectedIndex == 0 && comboBox3.SelectedIndex > 0)  //vyrobok
                    {
                        dtbl.DefaultView.RowFilter = string.Format("Vyrobok = '{0}'", comboBox3.SelectedItem.ToString());

                        comboBox1.Items.Clear();  //aktualizuje oddelenia
                        comboBox1.Items.Add("-Všetky-");
                        try
                        {
                            using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                            {
                                SqlCommand sqlCmd = new SqlCommand("SELECT Oddelenie FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Vyrobok ='" + comboBox3.SelectedItem + "'", sqlConnection);
                                sqlConnection.Open();
                                SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                                while (sqlReader.Read())
                                {
                                    comboBox1.Items.Add(sqlReader["Oddelenie"].ToString());
                                }
                                sqlReader.Close();
                            }
                        }
                        catch (Exception)
                        { }
                        object[] distinctItems1 = (from object o in comboBox1.Items select o).Distinct().ToArray();
                        comboBox1.Items.Clear();
                        comboBox1.Items.AddRange(distinctItems1);
                        comboBox1.SelectedIndex = 0;

                        comboBox2.Items.Clear();  //aktualizuje projekty
                        comboBox2.Items.Add("-Všetky-");
                        try
                        {
                            using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                            {
                                SqlCommand sqlCmd = new SqlCommand("SELECT Projekt FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Vyrobok ='" + comboBox3.SelectedItem + "'", sqlConnection);
                                sqlConnection.Open();
                                SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                                while (sqlReader.Read())
                                {
                                    comboBox2.Items.Add(sqlReader["Projekt"].ToString());
                                }
                                sqlReader.Close();
                            }
                        }
                        catch (Exception)
                        { }
                        object[] distinctItems2 = (from object o in comboBox2.Items select o).Distinct().ToArray();
                        comboBox2.Items.Clear();
                        comboBox2.Items.AddRange(distinctItems2);
                        comboBox2.SelectedIndex = 0;
                    }

                    if (comboBox1.SelectedIndex > 0 && comboBox2.SelectedIndex > 0 && comboBox3.SelectedIndex == 0)  //oddelenie + projekt
                    {
                        dtbl.DefaultView.RowFilter = string.Format("Oddelenie = '{0}' AND Projekt = '{1}' AND Aktiviny1Vymazany0 =  '{2}'", comboBox1.SelectedItem.ToString(), comboBox2.SelectedItem.ToString(),"1");

                        comboBox3.Items.Clear();  //aktualizuje vyrobky
                        comboBox3.Items.Add("-Všetky-");
                        try
                        {
                            using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                            {
                                SqlCommand sqlCmd = new SqlCommand("SELECT Vyrobok FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Oddelenie ='" + comboBox1.SelectedItem + "' AND Projekt = '" + comboBox2.SelectedItem +"'", sqlConnection);
                                sqlConnection.Open();
                                SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                                while (sqlReader.Read())
                                {
                                    comboBox3.Items.Add(sqlReader["Vyrobok"].ToString());
                                }
                                sqlReader.Close();
                            }
                        }
                        catch (Exception)
                        { }
                        object[] distinctItems3 = (from object o in comboBox3.Items select o).Distinct().ToArray();
                        comboBox3.Items.Clear();
                        comboBox3.Items.AddRange(distinctItems3);
                        comboBox3.SelectedIndex = 0;
                    }
                    if (comboBox1.SelectedIndex > 0 && comboBox2.SelectedIndex == 0 && comboBox3.SelectedIndex > 0)  //oddelenie + vyrobok
                    {
                        dtbl.DefaultView.RowFilter = string.Format("Oddelenie = '{0}' AND Vyrobok = '{1}' AND Aktiviny1Vymazany0 = '{2}'", comboBox1.SelectedItem.ToString(), comboBox3.SelectedItem.ToString(),"1");

                        comboBox2.Items.Clear();  //aktualizuje projekty
                        comboBox2.Items.Add("-Všetky-");
                        try
                        {
                            using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                            {
                                SqlCommand sqlCmd = new SqlCommand("SELECT Projekt FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Oddelenie ='" + comboBox1.SelectedItem + "' AND Vyrobok = '" + comboBox3.SelectedItem + "'", sqlConnection);
                                sqlConnection.Open();
                                SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                                while (sqlReader.Read())
                                {
                                    comboBox2.Items.Add(sqlReader["Projekt"].ToString());
                                }
                                sqlReader.Close();
                            }
                        }
                        catch (Exception)
                        { }
                        object[] distinctItems2 = (from object o in comboBox2.Items select o).Distinct().ToArray();
                        comboBox2.Items.Clear();
                        comboBox2.Items.AddRange(distinctItems2);
                        comboBox2.SelectedIndex = 0;
                    }
                    if (comboBox1.SelectedIndex == 0 && comboBox2.SelectedIndex > 0 && comboBox3.SelectedIndex > 0)  //projekt + vyrobok
                    {
                        dtbl.DefaultView.RowFilter = string.Format("Projekt = '{0}' AND Vyrobok = '{1}' AND Aktiviny1Vymazany0 =  '{2}'", comboBox2.SelectedItem.ToString(), comboBox3.SelectedItem.ToString(),"1");

                        comboBox1.Items.Clear();  //aktualizuje oddelenia
                        comboBox1.Items.Add("-Všetky-");
                        try
                        {
                            using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                            {
                                SqlCommand sqlCmd = new SqlCommand("SELECT Oddelenie FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Projekt ='" + comboBox2.SelectedItem + "' AND Vyrobok = '" + comboBox3.SelectedItem + "'", sqlConnection);
                                sqlConnection.Open();
                                SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                                while (sqlReader.Read())
                                {
                                    comboBox1.Items.Add(sqlReader["Oddelenie"].ToString());
                                }
                                sqlReader.Close();
                            }
                        }
                        catch (Exception)
                        { }
                        object[] distinctItems1 = (from object o in comboBox1.Items select o).Distinct().ToArray();
                        comboBox1.Items.Clear();
                        comboBox1.Items.AddRange(distinctItems1);
                        comboBox1.SelectedIndex = 0;
                    }

                    if (comboBox1.SelectedIndex > 0 && comboBox2.SelectedIndex > 0 && comboBox3.SelectedIndex > 0)  //oddelenie + projekt + vyrobok
                    {
                        dtbl.DefaultView.RowFilter = string.Format("Oddelenie = '{0}' AND Projekt = '{1}' AND Vyrobok = '{2}' AND Aktiviny1Vymazany0 =  '{3}'", comboBox1.SelectedItem.ToString(), comboBox2.SelectedItem.ToString(), comboBox3.SelectedItem.ToString(),"1");
                    }
                    dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;
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
                    try  
                    {
                        sqlCon.Open();
                        DataGridViewRow dgvRow = dataGridView1.CurrentRow;
                        SqlCommand sqlCmd = new SqlCommand("SM30EDIT", sqlCon);
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.Parameters.AddWithValue("@ID", dgvRow.Cells["txtID"].Value == DBNull.Value ? "" : dgvRow.Cells["txtID"].Value.ToString());
                        sqlCmd.Parameters.AddWithValue("@Oddelenie", dgvRow.Cells["txtOddelenie"].Value == DBNull.Value ? "" : dgvRow.Cells["txtOddelenie"].Value.ToString());
                        sqlCmd.Parameters.AddWithValue("@Projekt", dgvRow.Cells["txtProjekt"].Value == DBNull.Value ? "" : dgvRow.Cells["txtProjekt"].Value.ToString());
                        sqlCmd.Parameters.AddWithValue("@Vyrobok", dgvRow.Cells["txtVyrobok"].Value == DBNull.Value ? "" : dgvRow.Cells["txtVyrobok"].Value.ToString());
                        sqlCmd.Parameters.AddWithValue("@Stanica", dgvRow.Cells["txtStanica"].Value == DBNull.Value ? "" : dgvRow.Cells["txtStanica"].Value.ToString());
                        sqlCmd.Parameters.AddWithValue("@AlternativnaStanica", dgvRow.Cells["txtAlternativnaStanica"].Value == DBNull.Value ? "" : dgvRow.Cells["txtAlternativnaStanica"].Value.ToString());
                        sqlCmd.Parameters.AddWithValue("@SingleAleboGroup", dgvRow.Cells["txtSingleAleboGroup"].Value == DBNull.Value ? "" : dgvRow.Cells["txtSingleAleboGroup"].Value.ToString());
                        sqlCmd.Parameters.AddWithValue("@Material", dgvRow.Cells["txtMaterial"].Value == DBNull.Value ? "" : dgvRow.Cells["txtMaterial"].Value.ToString());
                        sqlCmd.Parameters.AddWithValue("@AlternativnyMaterial", dgvRow.Cells["txtAlternativnyMaterial"].Value == DBNull.Value ? "" : dgvRow.Cells["txtAlternativnyMaterial"].Value.ToString());
                        sqlCmd.Parameters.AddWithValue("@Popis", dgvRow.Cells["txtPopis"].Value == DBNull.Value ? "" : dgvRow.Cells["txtPopis"].Value.ToString());
                        sqlCmd.Parameters.AddWithValue("@PopisHV", dgvRow.Cells["txtPopisHV"].Value == DBNull.Value ? "" : dgvRow.Cells["txtPopisHV"].Value.ToString());
                        sqlCmd.Parameters.AddWithValue("@Aktiviny1Vymazany0", dgvRow.Cells["txtAktiviny1Vymazany0"].Value == DBNull.Value ? "" : dgvRow.Cells["txtAktiviny1Vymazany0"].Value.ToString());
                        sqlCmd.Parameters.AddWithValue("@PridalUpravil", Environment.MachineName + " & " + Environment.UserName);
                        sqlCmd.Parameters.AddWithValue("@DatumPridaniaUpravy", DateTime.Now);
                        sqlCmd.ExecuteNonQuery();
                        //PopulateDataGridView();
                    }
                    catch (Exception E)
                    { MessageBox.Show(E.ToString()); }
                }
            }
        }

        private void dataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            try
            {
                if (dataGridView1.CurrentRow.Cells["txtID"].Value != DBNull.Value)
                {
                    if (MessageBox.Show("Vymazať označený riadok?", "VÝMAZ", MessageBoxButtons.YesNo) == DialogResult.Yes)
                    {
                        foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                        {
                            using (SqlConnection sqlCon = new SqlConnection(Connection.ConnectionString))
                            {
                                sqlCon.Open();
                                SqlCommand sqlCmd = new SqlCommand("SM30DELETE", sqlCon);
                                sqlCmd.CommandType = CommandType.StoredProcedure;
                                sqlCmd.Parameters.AddWithValue("@ID", (int)row.Cells["txtID"].Value);
                                sqlCmd.Parameters.AddWithValue("@Vymazal", Environment.MachineName + " & " + Environment.UserName);
                                sqlCmd.Parameters.AddWithValue("@DatumVymazu", DateTime.Now);
                                sqlCmd.ExecuteNonQuery();
                            }
                        }
                    }
                    else
                        e.Cancel = true;
                }
                else
                    e.Cancel = true;
            }
            catch (Exception E)
            { MessageBox.Show(E.ToString()); }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {   //+ reference Microsoft.office.interop.Excel + using.., required + EditColumns -> SortMode -> Programmatic for each column
                copyAlltoClipboard();
                Excel.Application xlexcel;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;
                xlexcel = new Excel.Application();
                xlexcel.Visible = true;
                xlWorkBook = xlexcel.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                //najprv naformátuje založku na text aby sa nestratili nuly na zaciatku v nazve materiálu
                Excel.Range cells = xlWorkBook.Worksheets[1].Cells;  
                cells.NumberFormat = "@";

                Excel.Range CR = (Excel.Range)xlWorkSheet.Cells[1, 1];
                CR.Select();
                xlWorkSheet.PasteSpecial(CR, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
            }
            catch (Exception) { MessageBox.Show("Excel nie je podporovaný na tomto PC"); }
        }
        private void copyAlltoClipboard()
        {
                dataGridView1.SelectionMode = DataGridViewSelectionMode.FullColumnSelect;
                dataGridView1.SelectAll();
                DataObject dataObj = dataGridView1.GetClipboardContent();
                if (dataObj != null)
                Clipboard.SetDataObject(dataObj);
                dataGridView1.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
        }
        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            PopulateDataGridView();
        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            PopulateDataGridView();
        }

        private void comboBox2_SelectionChangeCommitted(object sender, EventArgs e)
        {
            PopulateDataGridView();
        }

        private void comboBox3_SelectionChangeCommitted(object sender, EventArgs e)
        {
            PopulateDataGridView();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Vymazať všetky dáta natrvalo?", "VÝMAZ", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                {
                    SqlCommand sqlCmd = new SqlCommand("DELETE FROM SM30", sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                    sqlReader.Close();
                }
                PopulateDataGridView();
            }
            else
            {
             return;
            }
        }

        private void button4_Click(object sender, EventArgs e)  //nájde duplicity, ktoré majú rovnakú hodnotu v stĺpcoch: projekt, výrobok, stanica a materiál
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                Settings.Default.StatusDuplicity = "0";
                Settings.Default.Save();
                for (int i = 0; i < dataGridView1.RowCount -1; i++)
                {
                    if (dataGridView1.Rows[i].Cells[7].Value.ToString() != "" && dataGridView1.Rows[i].Cells[8].Value.ToString() != "" && dataGridView1.Rows[i].Cells[9].Value.ToString() != "" && dataGridView1.Rows[i].Cells[12].Value.ToString() != "" && dataGridView1.Rows[i].Cells[16].Value.ToString() == "1")
                    {
                        string overovany = dataGridView1.Rows[i].Cells[7].Value.ToString() + dataGridView1.Rows[i].Cells[8].Value.ToString() + dataGridView1.Rows[i].Cells[9].Value.ToString() + dataGridView1.Rows[i].Cells[12].Value.ToString();
                        for (int j = i + 1; j < dataGridView1.RowCount - 1; j++)
                        {
                            string porovnavany = dataGridView1.Rows[j].Cells[7].Value.ToString() + dataGridView1.Rows[j].Cells[8].Value.ToString() + dataGridView1.Rows[j].Cells[9].Value.ToString() + dataGridView1.Rows[j].Cells[12].Value.ToString();
                            if (overovany == porovnavany)
                            {
                                dataGridView1.Rows[j].Selected = true;  //označí duplicity
                                dataGridView1.Rows[i].Selected = true;
                                MessageBox.Show("Nájdená duplicita: " +
                                 Environment.NewLine + "Projekt: " + dataGridView1.Rows[i].Cells[7].Value.ToString() +
                                 Environment.NewLine + "Výrobok: " + dataGridView1.Rows[i].Cells[8].Value.ToString() +
                                 Environment.NewLine + "Stanica: " + dataGridView1.Rows[i].Cells[9].Value.ToString() +
                                 Environment.NewLine + "Materiál: " + dataGridView1.Rows[i].Cells[12].Value.ToString());
                                Settings.Default.StatusDuplicity = "1";
                                Settings.Default.Save();
                            }
                        }
                    }
                }
                if (Settings.Default.StatusDuplicity == "0")
                {
                    MessageBox.Show("Žiadne duplicity");
                }
            }
            catch(Exception W)
            {MessageBox.Show(W.ToString());}
            Cursor = Cursors.Default;
        }
    }
}
