using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;
using System.Globalization;

namespace Waste_registration
{
    public partial class Nastavenia_VstupneData : Form
    {
        public Nastavenia_VstupneData()
        {
            InitializeComponent();
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Nastavenia nextForm = new Nastavenia();
            Hide();
            nextForm.ShowDialog();
            Close();
        }
        private void Nastavenia_VstupneData_Load(object sender, EventArgs e)
        {
            PopulateDataGridView();
        }
        void PopulateDataGridView()
        {
            using (SqlConnection sqlCon = new SqlConnection(Connection.ConnectionString))
            {
                try
                {
                    //sqlCon.Open();
                    //SqlDataAdapter sqlDa = new SqlDataAdapter("SELECT * FROM SM30", sqlCon);
                    //DataTable dtbl = new DataTable();
                    //sqlDa.Fill(dtbl);
                    //dataGridView1.DataSource = dtbl;
                    //dtbl.DefaultView.RowFilter = string.Format("Status LIKE '{0}'", "1");  //Zobrazí len aktívne/nezmazané riadky
                }
                catch (Exception) { }
            }
        }



        private void button3_Click(object sender, EventArgs e)
        {
            //dataGridView1.AutoGenerateColumns = false;  //
            //dataGridView1.DataSource = null;

            try
            {
                //    string s = Clipboard.GetText();
                //    string[] lines = s.Replace("\n", "").Split('\r');
                //    dataGridView1.Rows.Add(lines.Length - 1);
                //    string[] fields;

                //    int row = dataGridView1.Rows.Count -1;//
                //    int col = 1;
                //    foreach (string item in lines)
                //    {
                //        fields = item.Split('\t');
                //        foreach (string f in fields)
                //        {
                //            dataGridView1[col, row].Value = f;
                //            col++;
                //        }
                //        row++;   //riadok od 0 do konca označeného riadka
                //        col = 1;  //počiatočný stĺpec
                //        { }
                //    }

                //MessageBox.Show(dataGridView1.Rows.Count.ToString());






                string s = Clipboard.GetText();
                string[] lines = s.Replace("\n", "").Split('\r');
                dataGridView1.Rows.Add(lines.Length );
                string[] fields;
                int row = 0;
                int col = 1;
                foreach (string item in lines)
                {
                    fields = item.Split('\t');
                    foreach (string f in fields)
                    {
                        dataGridView1[col, row].Value = f;
                        col++;
                    }
                    row++;   
                    col = 1;
                    { }
                }

                button4.Visible = true;
                button3.Visible = false;
                button5.Visible = true;






                ////vymaže prázdne riadky
                //for (int i = 1; i < dataGridView1.RowCount - 1; i++)
                //{
                //    if (dataGridView1.Rows[i].Cells[0].Value.ToString().Trim() == "" || dataGridView1.Rows[i].Cells[1].Value.ToString().Trim() == "")
                //    {
                //        dataGridView1.Rows.RemoveAt(i);
                //        i--;
                //    }
                //}


                //                if (dataGridView1 == null || dataGridView1.ToString() == "")
                //{
                //    MessageBox.Show("Žiadne dáta neboli vložené");
                //    dataGridView1.Rows.Clear();
                //    return;
                //}
                //else
                //{
                //    button4.Visible = true;
                //    button3.Visible = false;
                //}


                ////overiť vstupné dáta


                //rozdeliť na hromadné pridanie dát a úpravu po jednom do dvoch okien!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

                //////////////////////////////////////////////////////////////nahraj do SQL


                //string s = Clipboard.GetText();
                //string[] lines = s.Replace("\n", "").Split('\r');
                //dataGridView1.Rows.Add(lines.Length - 1);
                //string[] fields;
                //int row = 0;
                //int col = 0;
                //foreach (string item in lines)
                //{
                //    fields = item.Split('\t');
                //    foreach (string f in fields)
                //    {
                //        dataGridView1[col, row].Value = f;
                //        col++;
                //    }
                //    row++;   //riadok od 0 do konca označeného riadka
                //    col = 1;  //počiatočný stĺpec
                //    { }
                //}
            }
            catch (Exception) 
            { 
                MessageBox.Show("Nekompletné dáta! Overte ich a opakujte vloženie.");
                dataGridView1.Rows.Clear();
            }


            //foreach (DataGridViewRow dr in dataGridView1.Rows)
            //{
            //    foreach (DataGridViewCell dc in dr.Cells)
            //    {
            //        if (dc.Value == "")
            //        {
            //            MessageBox.Show("Žiadne dáta neboli vložené");
            //            dataGridView1.Rows.Clear();
            //            return;
            //        }
            //        if (dc.Value.ToString() != "" || dc.Value.ToString() != null)
            //        {
            //            button4.Visible = true;
            //            button3.Visible = false;
            //        }
            //    }
            //}



            //try
            //{
            //    string s = Clipboard.GetText();
            //    string[] lines = s.Replace("\n", "").Split('\r');

            //        dataGridView1.Rows.Add(lines.Length);

            //        string[] fields;
            //        int row = dataGridView1.CurrentRow.Index;  //vkladá od označeného riadku
            //                                                   //int row = 0; //riadok od 0 do konca označeného riadka
            //        int col = 1; //počiatočný stĺpec (vynecháva ID-skrytý stĺpec)
            //        foreach (string item in lines)
            //        {
            //            fields = item.Split('\t');
            //            foreach (string f in fields)
            //            {
            //                dataGridView1[col, row].Value = f;
            //                col++;
            //            }
            //            row++;
            //            col = 1;  //počiatočný stĺpec (vynecháva ID-skrytý stĺpec)
            //            { }
            //        }
            //    }
            //catch (Exception E) { MessageBox.Show(E.ToString()); }

        }
        

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)  ////pridať overenie správnosti údajov pred zápisom do SQL!!!!!!!
        {
            if (dataGridView1.CurrentRow != null)
            {

                    using (SqlConnection sqlCon = new SqlConnection(Connection.ConnectionString))
                    {
                    try   //musí byť všetko splnené aby vykonalo komplet zápis
                    {
                        //sqlCon.Open();
                        //DataGridViewRow dgvRow = dataGridView1.CurrentRow;
                        //SqlCommand sqlCmd = new SqlCommand("SM30EDIT", sqlCon);
                        //sqlCmd.CommandType = CommandType.StoredProcedure;
                        //sqlCmd.Parameters.AddWithValue("@ID", dgvRow.Cells["txtID"].Value == DBNull.Value ? "" : dgvRow.Cells["txtID"].Value.ToString());
                        //sqlCmd.Parameters.AddWithValue("@Oddelenie", dgvRow.Cells["txtOddelenie"].Value == DBNull.Value ? "" : dgvRow.Cells["txtOddelenie"].Value.ToString());
                        //sqlCmd.Parameters.AddWithValue("@Vyrobok", dgvRow.Cells["txtVyrobok"].Value == DBNull.Value ? "" : dgvRow.Cells["txtVyrobok"].Value.ToString());
                        //sqlCmd.Parameters.AddWithValue("@Stanica", dgvRow.Cells["txtStanica"].Value == DBNull.Value ? "" : dgvRow.Cells["txtStanica"].Value.ToString());
                        //sqlCmd.Parameters.AddWithValue("@SingleAleboGroup", dgvRow.Cells["txtSingleAleboGroup"].Value == DBNull.Value ? "" : dgvRow.Cells["txtSingleAleboGroup"].Value.ToString());
                        //sqlCmd.Parameters.AddWithValue("@Material", dgvRow.Cells["txtMaterial"].Value == DBNull.Value ? "" : dgvRow.Cells["txtMaterial"].Value.ToString());

                        //sqlCmd.Parameters.AddWithValue("@KsMaterialuNa1VyrobokStanicu", dgvRow.Cells["txtKsMaterialuNa1VyrobokStanicu"].Value == DBNull.Value ? "" : dgvRow.Cells["txtKsMaterialuNa1VyrobokStanicu"].Value.ToString());

                        //sqlCmd.Parameters.AddWithValue("@MiestoOdpisu", dgvRow.Cells["txtMiestoOdpisu"].Value == DBNull.Value ? "" : dgvRow.Cells["txtMiestoOdpisu"].Value.ToString());
                        //sqlCmd.Parameters.AddWithValue("@MinutovaMzda", dgvRow.Cells["txtMinutovaMzda"].Value == DBNull.Value ? "" : dgvRow.Cells["txtMinutovaMzda"].Value.ToString());
                        //sqlCmd.Parameters.AddWithValue("@CA03LaborDelenoZaklMnozstvoVmin", dgvRow.Cells["txtCA03LaborDelenoZaklMnozstvoVmin"].Value == DBNull.Value ? "" : dgvRow.Cells["txtCA03LaborDelenoZaklMnozstvoVmin"].Value.ToString());
                        //if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString() != "" && dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value.ToString() != "")
                        //{
                        //    string MinNaKus = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString();
                        //    string EurNaMin = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value.ToString();
                        //    double EurNaKus = Convert.ToDouble(MinNaKus) * Convert.ToDouble(EurNaMin);
                        //    double EurNaKusFormat = Convert.ToDouble(EurNaKus, CultureInfo.InvariantCulture);
                        //    sqlCmd.Parameters.AddWithValue("@WageFinal", EurNaKusFormat.ToString());
                        //}

                        //sqlCmd.Parameters.AddWithValue("@Aktiviny1Vymazany0", "1");
                        //sqlCmd.Parameters.AddWithValue("@PridalUpravil", Environment.MachineName);
                        //sqlCmd.Parameters.AddWithValue("@DatumPridaniaUpravy", DateTime.Now);


                        //sqlCmd.ExecuteNonQuery();
                        //PopulateDataGridView();
                    }
                    catch (Exception)
                    {}
                }
                //sqlCmd.Parameters.AddWithValue("@WageFinal", dgvRow.Cells["txtWageFinal"].Value == DBNull.Value ? "" : dgvRow.Cells["txtWageFinal"].Value.ToString());
                //Hodnota odpadu - práca =  CA03LaborDelenoZaklMnozstvoVmin * MinutovaMzda

                //string MinNaKus = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString();
                //string EurNaMin = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[10].Value.ToString();
                //double EurNaKus = Convert.ToDouble(MinNaKus) * Convert.ToDouble(EurNaMin);
                //double EurNaKusFormat = Convert.ToDouble(EurNaKus, System.Globalization.CultureInfo.InvariantCulture);
                //sqlCmd.Parameters.AddWithValue("@WageFinal", EurNaKusFormat.ToString());
                //if (dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[6].Value.ToString() != "")
                //{
                //    //string MinNaKus = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[9].Value.ToString();
                //    //string EurNaMin = dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value.ToString();
                //    //double EurNaKus = Convert.ToDouble(MinNaKus) * Convert.ToDouble(EurNaMin);
                //}

                //sqlCmd.Parameters.AddWithValue("KsMaterialuNa1VyrobokStanicuKsMaterialuNa1VyrobokStanicu", Mnozstvo.ToString());

                //double Mnozstvo = Convert.ToDouble(dataGridView1.Rows[dataGridView1.CurrentCell.RowIndex].Cells[8].Value, CultureInfo.InvariantCulture);


            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(Connection.ConnectionString))
            {
                using (SqlCommand cmd = new SqlCommand("INSERT INTO SM30 (Oddelenie, Vyrobok, SingleAleboGroup) VALUES (@Oddelenie, @Vyrobok, @SingleAleboGroup)", con))
                {
                    cmd.Parameters.Add(new SqlParameter("@Oddelenie", SqlDbType.VarChar));
                    cmd.Parameters.Add(new SqlParameter("@Vyrobok", SqlDbType.VarChar));
                    cmd.Parameters.Add(new SqlParameter("@SingleAleboGroup", SqlDbType.VarChar));
                    con.Open();
                    foreach (DataGridViewRow roww in dataGridView1.Rows)
                    {
                        if (!roww.IsNewRow)
                        {
                            cmd.Parameters["@Oddelenie"].Value = roww.Cells[1].Value;
                            cmd.Parameters["@Vyrobok"].Value = roww.Cells[2].Value; 
                            cmd.Parameters["@SingleAleboGroup"].Value = roww.Cells[2].Value;
                            cmd.ExecuteNonQuery();
                            }
                    }
                }
            }

                //while (dataGridView1.Rows.Count > 0)
                //{
                //    dataGridView1.Rows.RemoveAt(0);
                //}
                
                //button3.Visible = true;
            }
            catch (Exception) {}
            MessageBox.Show("Importované dáta boli nahrané do databázy");
            dataGridView1.Rows.Clear();
            button4.Visible = false;
            button5.Visible = false;

        }

        private void button5_Click(object sender, EventArgs e)
        {
            Nastavenia_VstupneData nextForm = new Nastavenia_VstupneData();
            Hide();
            nextForm.ShowDialog();
            Close();
        }
    }
}
