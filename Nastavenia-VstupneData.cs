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
            try  //počiatočný ping serveru
            {
                using (SqlConnection connection = new SqlConnection(Connection.ConnectionString))
                    connection.Open();
            }
            catch (Exception)
            {
                MessageBox.Show("Spojenie so serverom zlyhalo, dáta nemožno zobraziť");
                button3.Visible = false;
                return;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Nastavenia nextForm = new Nastavenia();
            Hide();
            nextForm.ShowDialog();
            Close();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
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
            }
            catch (Exception) 
            { 
                MessageBox.Show("Nekompletné dáta! Overte ich a opakujte vloženie.");
                dataGridView1.Rows.Clear();
            }
        }
        private void button4_Click(object sender, EventArgs e) //Nahraj do databázy
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                using (SqlConnection con = new SqlConnection(Connection.ConnectionString))
                {
                    using (SqlCommand cmd = new SqlCommand("INSERT INTO SM30 (Oddelenie,Projekt,Vyrobok,Stanica,AlternativnaStanica,SingleAleboGroup,Material,AlternativnyMaterial,Popis,PopisHV,Aktiviny1Vymazany0,PridalUpravil,DatumPridaniaUpravy) VALUES (@Oddelenie,@Projekt,@Vyrobok,@Stanica,@AlternativnaStanica,@SingleAleboGroup,@Material,@AlternativnyMaterial,@Popis,@PopisHV,@Aktiviny1Vymazany0,@PridalUpravil,@DatumPridaniaUpravy)", con))
                    {
                        cmd.Parameters.Add(new SqlParameter("@Oddelenie", SqlDbType.VarChar));
                        cmd.Parameters.Add(new SqlParameter("@Projekt", SqlDbType.VarChar));
                        cmd.Parameters.Add(new SqlParameter("@Vyrobok", SqlDbType.VarChar));
                        cmd.Parameters.Add(new SqlParameter("@Stanica", SqlDbType.VarChar));
                        cmd.Parameters.Add(new SqlParameter("@AlternativnaStanica", SqlDbType.VarChar));
                        cmd.Parameters.Add(new SqlParameter("@SingleAleboGroup", SqlDbType.VarChar));
                        cmd.Parameters.Add(new SqlParameter("@Material", SqlDbType.VarChar));
                        cmd.Parameters.Add(new SqlParameter("@AlternativnyMaterial", SqlDbType.VarChar));
                        cmd.Parameters.Add(new SqlParameter("@Popis", SqlDbType.VarChar));
                        cmd.Parameters.Add(new SqlParameter("@PopisHV", SqlDbType.VarChar));
                        cmd.Parameters.Add(new SqlParameter("@Aktiviny1Vymazany0", SqlDbType.VarChar));
                        cmd.Parameters.Add(new SqlParameter("@PridalUpravil", SqlDbType.VarChar));
                        cmd.Parameters.Add(new SqlParameter("@DatumPridaniaUpravy", SqlDbType.DateTime));
                        con.Open();
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (!row.IsNewRow)
                            {
                                cmd.Parameters["@Oddelenie"].Value = row.Cells[1].Value;
                                cmd.Parameters["@Projekt"].Value = row.Cells[2].Value;
                                cmd.Parameters["@Vyrobok"].Value = row.Cells[3].Value;
                                cmd.Parameters["@Stanica"].Value = row.Cells[4].Value;
                                cmd.Parameters["@AlternativnaStanica"].Value = row.Cells[5].Value;
                                cmd.Parameters["@SingleAleboGroup"].Value = row.Cells[6].Value;
                                cmd.Parameters["@Material"].Value = row.Cells[7].Value;
                                cmd.Parameters["@AlternativnyMaterial"].Value = row.Cells[8].Value;
                                cmd.Parameters["@Popis"].Value = row.Cells[9].Value;
                                cmd.Parameters["@PopisHV"].Value = row.Cells[10].Value;
                                cmd.Parameters["@Aktiviny1Vymazany0"].Value = "1";
                                cmd.Parameters["@PridalUpravil"].Value = Environment.MachineName + " & " + Environment.UserName;
                                cmd.Parameters["@DatumPridaniaUpravy"].Value = DateTime.Now;
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }
            catch (Exception) { Cursor = Cursors.Default; }
            Cursor = Cursors.Default;
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
