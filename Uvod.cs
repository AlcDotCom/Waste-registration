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
    public partial class Uvod : Form
    {
        public Uvod()
        {
            InitializeComponent();
            Settings.Default.Aktualnyuzivatel = "";
            Settings.Default.AktualneOddelenie = "";
            Settings.Default.Save();

            try  //počiatočný ping serveru
            {
                using (SqlConnection connection = new SqlConnection(Connection.ConnectionString))
                    connection.Open();
            }
            catch (Exception)
            {
                MessageBox.Show("Spojenie so serverom zlyhalo, dáta nemožno zobraziť");
            }

            try   //načíta oddelenia
            {
                using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                {
                    SqlCommand sqlCmd = new SqlCommand("SELECT Oddelenie FROM SM30", sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        comboBox2.Items.Add(sqlReader["Oddelenie"].ToString());
                    }
                    sqlReader.Close();
                }
                object[] distinctItems2 = (from object o in comboBox2.Items select o).Distinct().ToArray();
                comboBox2.Items.Clear();
                comboBox2.Items.AddRange(distinctItems2);
                comboBox2.SelectedIndex = -1;
            }
            catch (Exception) {}

            int i = 0;
            while (comboBox2.Items.Count - 1 >= i)
            {
                if (Convert.ToString(comboBox2.Items[i]).Trim() == string.Empty)
                {
                    comboBox2.Items.RemoveAt(i);
                    i -= 1;
                }
                i += 1;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Dispose();
            heslo nextForm = new heslo();
            Hide();
            nextForm.ShowDialog();
            Close();
        }
        public new void Dispose()
        {
            pictureBox1.Image.Dispose();
            pictureBox1.Image = null;
        }

        private void cbxDesign_DrawItem(object sender, DrawItemEventArgs e)
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

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text.Length == 0)
            {
                MessageBox.Show("Vyberte oddelenie");
                comboBox2.Focus();
                return;
            }

            else
            {
                Settings.Default.Aktualnyuzivatel = Environment.UserName;
                Settings.Default.AktualneOddelenie = comboBox2.Text;
                Settings.Default.Save();

                dispose();
                VyberProjektu nextForm = new VyberProjektu();
                Hide();
                nextForm.ShowDialog();
                Close();
            }
        }
        void dispose()
        {
            pictureBox1.Image.Dispose();
            pictureBox1.Image = null;
        }
    }
}
