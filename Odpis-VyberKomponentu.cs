using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Waste_registration.Properties;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using ComboBox = System.Windows.Forms.ComboBox;
using TextBox = System.Windows.Forms.TextBox;

namespace Waste_registration
{
    public partial class _1 : Form
    {
        public static string duplicities;
        public static int A = 100;

        public _1()
        {
            InitializeComponent();
            textBox3.Text = 1.ToString();
            textBox3.ReadOnly = true;
            label3.Text = Settings.Default.Aktualnyprojekt;
            Settings.Default.AktualnyMaterial = "";
            Settings.Default.Save();
            checkBox1.Checked = false;
            Text = "Výber položky na odpis na oddelení " + Settings.Default.AktualneOddelenie + " pre projekt " + Settings.Default.Aktualnyprojekt;
            NacitajVyrobky();
            NacitajChyby();
        }

        private void button2_Click(object sender, EventArgs e)  //tlačidlo späť
        {
            if (textBox1.Text == "")
            {
                Dispose();
                VyberProjektu nextForm = new VyberProjektu();
                Hide();
                nextForm.ShowDialog();
                Close();
            }
            else
            {
                DialogResult dialogResult = MessageBox.Show("Pridané položky do odpisu sa stratia, pokračovať?", "Ukončiť bez odpisu", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    VyberProjektu nextForm = new VyberProjektu();
                    Hide();
                    nextForm.ShowDialog();
                    Close();
                }
                else if (dialogResult == DialogResult.No)
                { return; }
            }
        }
        public new void Dispose()
        {
            pictureBox3.Image.Dispose();
            pictureBox3.Image = null;
            pictureBox4.Image.Dispose();
            pictureBox4.Image = null;
        }
        private void cbxDesign_DrawItem(object sender, DrawItemEventArgs e) //formátovanie textu vo výberových oknách nastred
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

        ///////////////////////////////////////////////////////// pre cbx kde treba zvyrazniť duplicity ////////////////////////////////
        private void cbxDesign_DrawItemDuplicities(object sender, DrawItemEventArgs e) //text in middle + highlight duplicities
        {

            ComboBox cbx = sender as ComboBox;
            if (cbx != null)
            {
                e.DrawBackground();
                string[] arr = duplicities.Split(',');
                if (!arr.Distinct().Contains(e.Index.ToString()) && e.Index != -1) //single
                {
                    StringFormat sf = new StringFormat();
                    sf.LineAlignment = StringAlignment.Center;
                    sf.Alignment = StringAlignment.Center;
                    Brush brush = new SolidBrush(cbx.ForeColor);
                    if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                        brush = SystemBrushes.HighlightText;
                    e.Graphics.DrawString(cbx.Items[e.Index].ToString(), cbx.Font, brush, e.Bounds, sf);
                }
                if (arr.Distinct().Contains(e.Index.ToString()) && e.Index != -1)  //duplicity
                {
                    StringFormat sf = new StringFormat();
                    sf.LineAlignment = StringAlignment.Center;
                    sf.Alignment = StringAlignment.Center;
                    Brush brush = new SolidBrush(Color.Red);
                    if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
                        brush = SystemBrushes.HighlightText;
                    e.Graphics.DrawString(cbx.Items[e.Index].ToString(), cbx.Font, brush, e.Bounds, sf);
                }
            }
        }
        ///////////////////////////////////////////////////////// pre cbx kde treba zvyrazniť duplicity ////////////////////////////////

        private void pictureBox3_Click(object sender, EventArgs e)   // +1
        {
            if (Settings.Default.AktualneOddelenie == "Vstrekovanie") //vstrekovanie
            {

                try
                {
                    if (Convert.ToDecimal(textBox3.Text) >= 999)
                    {
                        MessageBox.Show("Množstvo nemôže byť väčšie ako 999");
                        return;
                    }
                    else
                    {
                        textBox3.ReadOnly = true;
                        textBox3.Text = (Convert.ToDecimal(textBox3.Text) + 1).ToString();
                        textBox3.TabStop = false;
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Nesprávny formát zadaného množstva");
                    textBox3.Focus();
                    return;
                }
            }
            else //montáž
            {
                if (Convert.ToInt32(textBox3.Text) >= 999)
                {
                    MessageBox.Show("Množstvo nemôže byť väčšie ako 999");
                    return;
                }
                else
                {
                    textBox3.ReadOnly = true;
                    textBox3.Text = (Convert.ToInt32(textBox3.Text) + 1).ToString();
                    textBox3.TabStop = false;
                }
            }
        }

        private void pictureBox4_Click(object sender, EventArgs e)   // -1
        {
            if (Settings.Default.AktualneOddelenie == "Vstrekovanie") //vstrekovanie
            {
                try
                {
                    if (Convert.ToDecimal(textBox3.Text) < 1)
                    {
                        MessageBox.Show("Množstvo nemôže byť menšie ako 0");
                        return;
                    }
                    else
                    {
                        textBox3.ReadOnly = true;
                        textBox3.Text = (Convert.ToDecimal(textBox3.Text) - 1).ToString();
                        textBox3.TabStop = false;
                    }
                }
                catch(Exception)
                {
                    MessageBox.Show("Nesprávny formát zadaného množstva");
                    textBox3.Focus();
                    return;
                }
            }
            else //montáž
            {
                if (Convert.ToInt32(textBox3.Text) < 2)
                {
                    MessageBox.Show("Množstvo nemôže byť menšie ako 1");
                    return;
                }
                else
                {
                    textBox3.ReadOnly = true;
                    textBox3.Text = (Convert.ToInt32(textBox3.Text) - 1).ToString();
                    textBox3.TabStop = false;
                }
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (Settings.Default.AktualneOddelenie != "Vstrekovanie") //pre montáž a iné hneď pri zadávaní overí či je hodnota číslovka (musí byť celé číslo)
            {
                try
                {
                    int.Parse(textBox3.Text);
                }
                catch (Exception)
                {
                    MessageBox.Show("Zadaná hodnota musí byť číslovka a iba celé číslo");
                    textBox3.Text = "1";
                    textBox3.SelectAll();
                }
                if (Convert.ToInt32(textBox3.Text) > 999)
                {
                    MessageBox.Show("Množstvo nemôže byť väčšie ako 999");
                    textBox3.Text = "1";
                    textBox3.SelectAll();
                }
            }
        }
        private void textBox3_Click(object sender, EventArgs e)
        {
            textBox3.ReadOnly = false;
            textBox3.SelectAll();
        }
        void ZistiZmenu()
        {
            var teraz = DateTime.Now;
            DateTime t1 = Convert.ToDateTime("06:00:00");
            DateTime t2 = Convert.ToDateTime("14:00:00");
            DateTime t3 = Convert.ToDateTime("22:00:00");
            if (teraz >= t1 && teraz < t2)
            { Settings.Default.AktualnaZmena = "1"; Settings.Default.Save(); }
            else if (teraz >= t2 && teraz < t3)
            { Settings.Default.AktualnaZmena = "2"; Settings.Default.Save(); }
            else if (teraz >= t3)
            { Settings.Default.AktualnaZmena = "3"; Settings.Default.Save(); }
            else if (teraz < t1)
            { Settings.Default.AktualnaZmena = "3"; Settings.Default.Save(); }
            else
            { Settings.Default.AktualnaZmena = "N/A"; Settings.Default.Save(); }
        }
        
        void NacitajVyrobky()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                {
                    SqlCommand sqlCmd = new SqlCommand("SELECT Vyrobok, PopisHV FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND PopisHV <> '' AND Projekt = '" + Settings.Default.Aktualnyprojekt + "' AND Oddelenie = '" + Settings.Default.AktualneOddelenie + "'   ", sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        comboBox4.Items.Add(sqlReader["Vyrobok"].ToString() + "  •  " + sqlReader["PopisHV"].ToString());
                        comboBox3.Items.Add(sqlReader["Vyrobok"].ToString() + "  •  " + sqlReader["PopisHV"].ToString()); //pomocny cbx pre check duplicít
                    }
                    sqlReader.Close();
                }
                object[] distinctItems4 = (from object o in comboBox4.Items select o).Distinct().ToArray();
                comboBox4.Items.Clear();
                comboBox4.Items.AddRange(distinctItems4);

                comboBox4.SelectedIndex = -1;
                int i = 0;
                while (comboBox4.Items.Count - 1 >= i)
                {
                    if (Convert.ToString(comboBox4.Items[i]).Trim() == string.Empty)
                    {
                        comboBox4.Items.RemoveAt(i);
                        i -= 1;
                    }
                    i += 1;
                }

                /////////////////////////////////////////////////// pre highlight duplicít /////////////////////////////
                object[] distinctItems3 = (from object o in comboBox3.Items select o).Distinct().ToArray();
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(distinctItems3);
                object[] distinctItems33 = (from object o in comboBox3.Items select o.ToString().Substring(0, o.ToString().IndexOf('/'))).ToArray();
                comboBox3.Items.Clear();
                comboBox3.Items.AddRange(distinctItems33);
                duplicities = "";

                foreach (var item in comboBox3.Items)
                {
                    if (comboBox3.Items.IndexOf(item) == A)
                    {
                        duplicities = comboBox3.Items.IndexOf(item).ToString() + "," + duplicities;
                    }
                    A = comboBox3.Items.IndexOf(item);
                }
                /////////////////////////////////////////////////// pre highlight duplicít /////////////////////////////

            }
            catch (Exception) {}
        }

        void NacitajKomponent()  //STANICA
        {
            Settings.Default.IndexVyberuStanica = comboBox1.SelectedIndex; //uloží vybratú položku
            Settings.Default.Save();

            comboBox1.Items.Clear();

            if (checkBox1.Checked)  //užívateľ vybral alternatívny odpis, načíta alternatívnu stanicu
            {
                try
                {
                    using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                    {
                        SqlCommand sqlCmd = new SqlCommand("SELECT AlternativnaStanica FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Projekt = '" + Settings.Default.Aktualnyprojekt + "' AND Oddelenie = '" + Settings.Default.AktualneOddelenie + "' AND Vyrobok = '" + Settings.Default.AktualnyVyrobok + "' ", sqlConnection);
                        sqlConnection.Open();
                        SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                        while (sqlReader.Read())
                        {
                            comboBox1.Items.Add(sqlReader["AlternativnaStanica"].ToString());
                        }
                        sqlReader.Close();
                    }
                    object[] distinctItems1 = (from object o in comboBox1.Items select o).Distinct().ToArray();
                    comboBox1.Items.Clear();
                    comboBox1.Items.AddRange(distinctItems1);
                    comboBox1.SelectedIndex = -1;
                    int i = 0;
                    while (comboBox1.Items.Count - 1 >= i)
                    {
                        if (Convert.ToString(comboBox1.Items[i]).Trim() == string.Empty)
                        {
                            comboBox1.Items.RemoveAt(i);
                            i -= 1;
                        }
                        i += 1;
                    }
                }
                catch (Exception) { }
            }
            else
            {
                try
                {
                    using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                    {
                        SqlCommand sqlCmd = new SqlCommand("SELECT Stanica FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Projekt = '" + Settings.Default.Aktualnyprojekt + "' AND Oddelenie = '" + Settings.Default.AktualneOddelenie + "' AND Vyrobok = '" + Settings.Default.AktualnyVyrobok + "' ", sqlConnection);
                        sqlConnection.Open();
                        SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                        while (sqlReader.Read())
                        {
                            comboBox1.Items.Add(sqlReader["Stanica"].ToString());
                        }
                        sqlReader.Close();
                    }
                    object[] distinctItems1 = (from object o in comboBox1.Items select o).Distinct().ToArray();
                    comboBox1.Items.Clear();
                    comboBox1.Items.AddRange(distinctItems1);
                    comboBox1.SelectedIndex = -1;
                    int i = 0;
                    while (comboBox1.Items.Count - 1 >= i)
                    {
                        if (Convert.ToString(comboBox1.Items[i]).Trim() == string.Empty)
                        {
                            comboBox1.Items.RemoveAt(i);
                            i -= 1;
                        }
                        i += 1;
                    }
                }
                catch (Exception) { }
            }
            comboBox1.SelectedIndex = Settings.Default.IndexVyberuStanica; //vyberie naposledy vybratú položku
        }

        void NacitajJednotlivy() 
        {
            comboBox2.Items.Clear();

            if (checkBox1.Checked)  //užívateľ vybral alternatívny odpis, načíta alternatívny materiál
            {
                try
                {
                    using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                    {
                        SqlCommand sqlCmd = new SqlCommand("SELECT AlternativnyMaterial, Popis FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND AlternativnaStanica = '" + comboBox1.SelectedItem.ToString() + "' AND Projekt = '" + Settings.Default.Aktualnyprojekt + "' AND Oddelenie = '" + Settings.Default.AktualneOddelenie + "' AND Vyrobok = '" + Settings.Default.AktualnyVyrobok + "' ", sqlConnection);
                        sqlConnection.Open();
                        SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                        while (sqlReader.Read())
                        {
                            comboBox2.Items.Add(sqlReader["AlternativnyMaterial"].ToString() + "  •  " + sqlReader["Popis"].ToString());
                        }
                        sqlReader.Close();
                    }
                    object[] distinctItems2 = (from object o in comboBox2.Items select o).Distinct().ToArray();
                    comboBox2.Items.Clear();
                    comboBox2.Items.AddRange(distinctItems2);
                    comboBox2.SelectedIndex = -1;
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
                catch (Exception) { }
            }
            else
            {
                try
                {
                    using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                    {
                        SqlCommand sqlCmd = new SqlCommand("SELECT Material, Popis FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Stanica = '" + comboBox1.SelectedItem.ToString() + "' AND Projekt = '" + Settings.Default.Aktualnyprojekt + "' AND Oddelenie = '" + Settings.Default.AktualneOddelenie + "' AND Vyrobok = '" + Settings.Default.AktualnyVyrobok + "' ", sqlConnection);
                        sqlConnection.Open();
                        SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                        while (sqlReader.Read())
                        {
                            comboBox2.Items.Add(sqlReader["Material"].ToString() + "  •  " + sqlReader["Popis"].ToString());
                        }
                        sqlReader.Close();
                    }
                    object[] distinctItems2 = (from object o in comboBox2.Items select o).Distinct().ToArray();
                    comboBox2.Items.Clear();
                    comboBox2.Items.AddRange(distinctItems2);
                    comboBox2.SelectedIndex = -1;
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
                catch (Exception) { }
            }
        }

        void NacitajVinnika()
        {
            comboBox6.Items.Clear();

            if (checkBox1.Checked)  //užívateľ vybral alternatívny odpis, načíta alternatívneho vinníka
            {
                try
                {
                    using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString)) 
                    {
                        SqlCommand sqlCmd = new SqlCommand("SELECT AlternativnyMaterial, Popis FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND AlternativnaStanica LIKE 'Jednotlivy material%' AND Projekt = '" + Settings.Default.Aktualnyprojekt + "' AND Oddelenie = '" + Settings.Default.AktualneOddelenie + "' AND Vyrobok = '" + Settings.Default.AktualnyVyrobok + "' ", sqlConnection);
                        sqlConnection.Open();
                        SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                        while (sqlReader.Read())
                        {
                            comboBox6.Items.Add(sqlReader["AlternativnyMaterial"].ToString() + "  •  " + sqlReader["Popis"].ToString());
                        }
                        sqlReader.Close();
                    }
                    object[] distinctItems6 = (from object o in comboBox6.Items select o).Distinct().ToArray();
                    comboBox6.Items.Clear();
                    comboBox6.Items.Add("NA");
                    comboBox6.Items.AddRange(distinctItems6);
                    comboBox6.SelectedIndex = -1;
                    int i = 0;
                    while (comboBox6.Items.Count - 1 >= i)
                    {
                        if (Convert.ToString(comboBox6.Items[i]).Trim() == string.Empty)
                        {
                            comboBox6.Items.RemoveAt(i);
                            i -= 1;
                        }
                        i += 1;
                    }
                }
                catch (Exception o) { MessageBox.Show(o.ToString()); }
            }
            else
            {
                try
                {
                    using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                    {
                        SqlCommand sqlCmd = new SqlCommand("SELECT Material, Popis FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Stanica = '" + comboBox1.Text.ToString() + "' AND Projekt = '" + Settings.Default.Aktualnyprojekt + "' AND Oddelenie = '" + Settings.Default.AktualneOddelenie + "' AND Vyrobok = '" + Settings.Default.AktualnyVyrobok + "' ", sqlConnection);
                        sqlConnection.Open();
                        SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                        while (sqlReader.Read())
                        {
                            comboBox6.Items.Add(sqlReader["Material"].ToString() + "  •  " + sqlReader["Popis"].ToString());
                        }
                        sqlReader.Close();
                    }
                    object[] distinctItems6 = (from object o in comboBox6.Items select o).Distinct().ToArray();
                    comboBox6.Items.Clear();
                    comboBox6.Items.Add("NA");
                    comboBox6.Items.AddRange(distinctItems6);
                    comboBox6.SelectedIndex = -1;
                    int i = 0;
                    while (comboBox6.Items.Count - 1 >= i)
                    {
                        if (Convert.ToString(comboBox6.Items[i]).Trim() == string.Empty)
                        {
                            comboBox6.Items.RemoveAt(i);
                            i -= 1;
                        }
                        i += 1;
                    }
                }
                catch (Exception) { }
            }
        }

        void NacitajChyby()
        {
            try
            {
                using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                {
                    SqlCommand sqlCmd = new SqlCommand("SELECT KodChyby, PopisChyby FROM SM30detail WHERE Oddelenie = '" + Settings.Default.AktualneOddelenie + "' AND KodChyby <> '' ", sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        comboBox5.Items.Add(sqlReader["KodChyby"].ToString() + "  •  " + sqlReader["PopisChyby"].ToString());
                    }
                    sqlReader.Close();
                }
                object[] distinctItems5 = (from object o in comboBox5.Items select o).Distinct().ToArray();
                comboBox5.Items.Clear();
                comboBox5.Items.AddRange(distinctItems5);
                comboBox5.SelectedIndex = -1;
                int i = 0;
                while (comboBox5.Items.Count - 1 >= i)
                {
                    if (Convert.ToString(comboBox5.Items[i]).Trim() == string.Empty)
                    {
                        comboBox5.Items.RemoveAt(i);
                        i -= 1;
                    }
                    i += 1;
                }
            }
            catch (Exception) {}
        }
        void ZistiAdresu()
        {
            try
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
            catch (Exception)
            {
                MessageBox.Show("Adresu uloženia nebolo možné načítať, overte pripojenie so serverom");
            }
        }

        private void button1_Click(object sender, EventArgs e)  //Pridanie komponentu na odpis
        {
            try
            {
                bool overenie = Convert.ToDecimal(textBox3.Text) <= 0;
                if (overenie)
                {
                    MessageBox.Show("Množstvo nemôže byť 0");
                    textBox3.Focus();
                    return;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("Nesprávny formát zadaného množstva");
                textBox3.Focus();
                return;
            }

            if (Settings.Default.AktualneOddelenie == "Vstrekovanie")
            {
                try  // pre vstrekovanie convertuje desatinnu hodnotu s ciarkou na bodku
                {
                    textBox3.Text = Convert.ToDecimal(textBox3.Text).ToString();
                    textBox3.Text = textBox3.Text.Replace(",", ".");
                }
                catch (Exception)
                {
                    MessageBox.Show("Množstvo nie je zadané správne");
                    textBox3.Focus();
                    return;
                }
            }

            if (comboBox4.Text.Length == 0)
            {
                MessageBox.Show("Vyberte výrobok");
                comboBox4.Focus();
                return;
            }
            else if (comboBox1.Text.Length == 0)
            {
                MessageBox.Show("Vyberte stanicu");
                comboBox1.Focus();
                return;
            }
            else
            {
                if (comboBox1.SelectedItem.ToString() == "Jednotlivy material")
                {
                    if (comboBox2.Text.Length == 0)
                    {
                        MessageBox.Show("Vyberte materiál");
                        comboBox2.Focus();
                        return;
                    }
                    else if (comboBox5.Text.Length == 0)
                    {
                        MessageBox.Show("Vyberte chybu");
                        comboBox5.Focus();
                        return;
                    }
                }
                else
                {
                    if (comboBox6.Text.Length == 0)
                    {
                        MessageBox.Show("Vyberte vinníka");
                        comboBox6.Focus();
                        return;
                    }
                    else if (comboBox5.Text.Length == 0)
                    {
                        MessageBox.Show("Vyberte chybu");
                        comboBox5.Focus();
                        return;
                    }
                    else if (comboBox2.Visible == true && comboBox2.Text.Length == 0)
                    {
                        MessageBox.Show("Vyberte materiál");
                        comboBox2.Focus();
                        return;
                    }
                }
            }
            ZistiZmenu();
            UlozChybu();

            if (textBox1.Text == "")
            {
                textBox1.Text += "Name,Project,date,time,Station,PN BB,quantity,shift,failure code,failure desc,line,caused by" + Environment.NewLine;
            }
            //                                    uzivatel                              vyrobok                                           datum                 
            string ItemToRegister = Settings.Default.Aktualnyuzivatel + "," + Settings.Default.AktualnyVyrobok + "," + DateTime.Now.ToString("dd") + "/" + DateTime.Now.ToString("MM") + "/" + DateTime.Now.ToString("yyyy") + "," +
            //                  čas                                          stanica                                   material                       množstvo
            string.Format("{0:HH:mm:ss}", DateTime.Now) + "," + comboBox1.SelectedItem.ToString() + "," + Settings.Default.AktualnyMaterial + "," + textBox3.Text + "," +
            //            zmena                          kod chyby                         popis chyby                                Linka
            Settings.Default.AktualnaZmena + "," + Settings.Default.KodChyby + "," + Settings.Default.PopisChyby + "," + Settings.Default.Aktualnyprojekt + "," +
            //        vinník              novy riadok
            Settings.Default.vinnik + Environment.NewLine;

            textBox1.Text += ItemToRegister;

            //comboBox4.SelectedIndex = -1;
            //comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            textBox3.Text = 1.ToString();
            comboBox6.Items.Clear();
            comboBox5.SelectedIndex = -1;

            if (textBox1.Text == "")    //update počtu položiek na odpis
            { label9.Visible = false;
                label9.Text = "";
            }
            else
            {
                label9.Text = "Počet položiek na odpis: " + (textBox1.Lines.Count()-2).ToString();
                label9.Visible = true;
                button3.Visible = true;
            }
            Settings.Default.AktualnyMaterial = "";
            Settings.Default.Save();

            NacitajKomponent();
            textBox2.Text = "";
            comboBox5.SelectedIndex = -1;
            NacitajStanicu();
        }

        private void button3_Click(object sender, EventArgs e)   //Odpis - uloženie do .csv
        {
            if (textBox1.Text == "") {}
            else
            {
                try
                {
                    if (Settings.Default.StatusAdresa == "1") //status určovaný v nastavení hesla
                    {
                        //uloženie na plochu užívateľa//
                        string adresa = string.Format(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)) + string.Format("/") + string.Format("{0:yyyyMMddHHmmssfff}", DateTime.Now) + ".csv";
                        string odpis = textBox1.Text;
                        File.AppendAllText(adresa, odpis);
                    }
                    else
                    { 
                        ZistiAdresu();
                        string adresa = string.Format(Settings.Default.AdresaOverenie) + string.Format("{0:yyyyMMddHHmmssfff}", DateTime.Now) + ".csv";
                        string odpis = textBox1.Text;
                        File.AppendAllText(adresa, odpis);
                    }
                    
                    textBox1.Clear();

                    var w = new Form() { Size = new Size(0, 0) };
                    Task.Delay(TimeSpan.FromSeconds(0.8))
                    .ContinueWith((t) => w.Close(), TaskScheduler.FromCurrentSynchronizationContext());
                    MessageBox.Show(w, "OK, odpis položiek prebehol úspešne.");

                    VyberProjektu nextForm = new VyberProjektu();
                    Hide();
                    nextForm.ShowDialog();
                    Close();
                }
                catch (Exception)
                { MessageBox.Show("Nie je možné odpísať zadané dáta, overte spojenie so serverom"); }
            }
        }

        private void label9_Click(object sender, EventArgs e)
        {
            if (textBox1.Visible == false)
            {
                int formWidth = Width;
                TextBox Mytextbox1 = new TextBox();
                Mytextbox1.Location = new Point(16, 15);
                textBox1.Size = new Size(formWidth - 40, 490);
                textBox1.Visible = true;
                textBox1.BringToFront();
            }
            else
            {
                TextBox Mytextbox1 = new TextBox();
                Mytextbox1.Location = new Point(16, 15);
                textBox1.Size = new Size(59, 59);
                textBox1.Visible = false;
            }
        }

        private void comboBox4_SelectionChangeCommitted(object sender, EventArgs e)  //výber HV
        {
            try
            {
                string[] FGwithDesc = comboBox4.SelectedItem.ToString().Split('•');
            Settings.Default.AktualnyVyrobok = FGwithDesc[0].Trim();
            Settings.Default.Save();
            checkBox1.Visible = false;

            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            comboBox6.Items.Clear();
            label2.Visible = false;
            comboBox2.Visible = false;
            NacitajKomponent();
            comboBox6.Items.Clear();
            textBox2.Text = "";
            comboBox5.SelectedIndex = -1;
            textBox3.Text = 1.ToString();

                //zisti ci dany PN má alternatívnu stanicu, ak áno checkbox na výber alt. odpisu bude viditeľný
                using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                {
                    SqlCommand sqlCmd = new SqlCommand("SELECT AlternativnaStanica FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Projekt = '" + Settings.Default.Aktualnyprojekt + "' AND Oddelenie = '" + Settings.Default.AktualneOddelenie + "' AND Vyrobok = '" + Settings.Default.AktualnyVyrobok + "' ", sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        string OverAlterStanicu = sqlReader["AlternativnaStanica"].ToString();
                        if (string.IsNullOrEmpty(OverAlterStanicu))
                        {
                            checkBox1.Visible = false;
                        }
                        else
                        {
                            checkBox1.Visible = true;
                        }
                    }
                    sqlReader.Close();
                }
            }
            catch (Exception)
            {} 


        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)  //výber stanice
        {
            NacitajStanicu();
        }

        void NacitajStanicu()
        {
            Settings.Default.AktualnyMaterial = ""; //zmaže vybraný materiál ak užívataľ prepne z jednotlivého na kompletný
            Settings.Default.Save();
            if (comboBox1.SelectedItem.ToString().Contains("Jednotlivy material"))  //Single
            {
                label2.Visible = true;
                comboBox2.Visible = true;
                comboBox6.Items.Clear();
                NacitajJednotlivy();
                label8.Visible = false;
                comboBox6.Visible = false;
            }
            else  //Groupy
            {
                string Stanica = comboBox1.SelectedItem.ToString();
                if (Stanica.Contains("'"))  ///////////////////////////////////////////ak je ' v texte stanice
                {
                    string StanicaBezQuote = Stanica.Replace("'", "•");
                    string[] StanicaLavaStrana = StanicaBezQuote.Split('•');
                    Settings.Default.VyhladavanaStanica = StanicaLavaStrana[0].Trim();
                    Settings.Default.Save();

                    if (checkBox1.Checked)
                    {
                        try
                        {
                            using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                            {
                                SqlCommand sqlCmd = new SqlCommand("SELECT SingleAleboGroup FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Vyrobok = '" + Settings.Default.AktualnyVyrobok + "'AND AlternativnaStanica LIKE '" + Settings.Default.VyhladavanaStanica + "%' ", sqlConnection);
                                sqlConnection.Open();
                                SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                                while (sqlReader.Read())
                                {
                                    string overenie = sqlReader["SingleAleboGroup"].ToString();
                                    if (overenie == "S")  ///Group odpisovaný jednotlivo
                                    {
                                        label2.Visible = true;
                                        comboBox2.Visible = true;
                                        NacitajVinnika();
                                        label8.Visible = true;
                                        comboBox6.Visible = true;
                                        NacitajJednotlivy();
                                    }
                                    else  ///Group odpisovaný štandardne
                                    {
                                        label2.Visible = false;
                                        comboBox2.Visible = false;
                                        NacitajVinnika();
                                        label8.Visible = true;
                                        comboBox6.Visible = true;
                                        comboBox2.Items.Clear();
                                        comboBox2.Items.Add("");
                                        comboBox2.SelectedIndex = 0;
                                    }
                                }
                                sqlReader.Close();
                            }
                        }
                        catch (Exception) { }
                    }
                    else
                    {
                        try
                        {
                            using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                            {
                                SqlCommand sqlCmd = new SqlCommand("SELECT SingleAleboGroup FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Vyrobok = '" + Settings.Default.AktualnyVyrobok + "' AND Stanica LIKE '" + Settings.Default.VyhladavanaStanica + "%' ", sqlConnection);
                                sqlConnection.Open();
                                SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                                while (sqlReader.Read())
                                {
                                    string overenie = sqlReader["SingleAleboGroup"].ToString();
                                    if (overenie == "S")  ///Group odpisovaný jednotlivo
                                    {
                                        label2.Visible = true;
                                        comboBox2.Visible = true;
                                        NacitajVinnika();
                                        label8.Visible = true;
                                        comboBox6.Visible = true;
                                        NacitajJednotlivy();
                                    }
                                    else  ///Group odpisovaný štandardne
                                    {
                                        label2.Visible = false;
                                        comboBox2.Visible = false;
                                        NacitajVinnika();
                                        label8.Visible = true;
                                        comboBox6.Visible = true;
                                        comboBox2.Items.Clear();
                                        comboBox2.Items.Add("");
                                        comboBox2.SelectedIndex = 0;
                                    }
                                }
                                sqlReader.Close();
                            }
                        }
                        catch (Exception) { }
                    }

                }
                else  //////////////////////////////// štandardne
                {
                    Settings.Default.VyhladavanaStanica = Stanica;
                    Settings.Default.Save();

                    if (checkBox1.Checked)
                    {
                        try
                        {
                            using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                            {
                                SqlCommand sqlCmd = new SqlCommand("SELECT SingleAleboGroup FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Vyrobok = '" + Settings.Default.AktualnyVyrobok + "'AND AlternativnaStanica = '" + Settings.Default.VyhladavanaStanica + "' ", sqlConnection);
                                sqlConnection.Open();
                                SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                                while (sqlReader.Read())
                                {
                                    string overenie = sqlReader["SingleAleboGroup"].ToString();
                                    if (overenie == "S")  ///Group odpisovaný jednotlivo
                                    {
                                        label2.Visible = true;
                                        comboBox2.Visible = true;
                                        NacitajVinnika();
                                        label8.Visible = true;
                                        comboBox6.Visible = true;
                                        NacitajJednotlivy();
                                    }
                                    else  ///Group odpisovaný štandardne
                                    {
                                        label2.Visible = false;
                                        comboBox2.Visible = false;
                                        NacitajVinnika();
                                        label8.Visible = true;
                                        comboBox6.Visible = true;
                                        comboBox2.Items.Clear();
                                        comboBox2.Items.Add("");
                                        comboBox2.SelectedIndex = 0;
                                    }
                                }
                                sqlReader.Close();
                            }
                        }
                        catch (Exception) { }
                    }
                    else
                    {
                        try
                        {
                            using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                            {
                                SqlCommand sqlCmd = new SqlCommand("SELECT SingleAleboGroup FROM SM30 WHERE Aktiviny1Vymazany0 = '1' AND Vyrobok = '" + Settings.Default.AktualnyVyrobok + "' AND Stanica = '" + Settings.Default.VyhladavanaStanica + "' ", sqlConnection);
                                sqlConnection.Open();
                                SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                                while (sqlReader.Read())
                                {
                                    string overenie = sqlReader["SingleAleboGroup"].ToString();
                                    if (overenie == "S")  ///Group odpisovaný jednotlivo
                                    {
                                        label2.Visible = true;
                                        comboBox2.Visible = true;
                                        NacitajVinnika();
                                        label8.Visible = true;
                                        comboBox6.Visible = true;
                                        NacitajJednotlivy();
                                    }
                                    else  ///Group odpisovaný štandardne
                                    {
                                        label2.Visible = false;
                                        comboBox2.Visible = false;
                                        NacitajVinnika();
                                        label8.Visible = true;
                                        comboBox6.Visible = true;
                                        comboBox2.Items.Clear();
                                        comboBox2.Items.Add("");
                                        comboBox2.SelectedIndex = 0;
                                    }
                                }
                                sqlReader.Close();
                            }
                        }
                        catch (Exception) { }
                    }
                }
            }

            if (Settings.Default.AktualneOddelenie == "Vstrekovanie") //pre vstrekovanie nebude na výber vinník, len NA v ponuke bez moznosti výberu
            {
                comboBox6.Items.Clear();
                comboBox6.Items.Add("NA");
                comboBox6.SelectedIndex = 0;
                Settings.Default.vinnik = comboBox6.SelectedItem.ToString(); ;
                Settings.Default.Save();
            }
            textBox2.Text = "";
            comboBox5.SelectedIndex = -1;
            textBox3.Text = 1.ToString();
        }

        private void comboBox2_SelectionChangeCommitted(object sender, EventArgs e)  //výber jednotlivého materiálu
        {
            string[] Material = comboBox2.SelectedItem.ToString().Split('•');  //uloží materiál na odpis
            Settings.Default.AktualnyMaterial = Material[0].Trim();
            Settings.Default.Save();

            if (Settings.Default.AktualneOddelenie != "Vstrekovanie")  //platí pre montáž
            {
                comboBox6.Items.Clear();
                comboBox6.Items.Add(Settings.Default.AktualnyMaterial);
                comboBox6.SelectedIndex = 0;
                Settings.Default.vinnik = comboBox6.SelectedItem.ToString(); //pre montáž bude vinník vždy rovnaký ako materiál
                Settings.Default.Save();
            }

            textBox2.Text = "";
            comboBox5.SelectedIndex = -1;
            textBox3.Text = 1.ToString();
        }
        private void comboBox5_SelectionChangeCommitted_1(object sender, EventArgs e)
        {
            UlozChybu();
        }
        void UlozChybu()
        {
            string[] KodChyby = comboBox5.SelectedItem.ToString().Split('•');  //uloží kód chyby na odpis
            Settings.Default.KodChyby = KodChyby[0].Trim();
            Settings.Default.Save();

            string[] PopisChyby = comboBox5.SelectedItem.ToString().Split('•');  //uloží popis chyby na odpis
            Settings.Default.PopisChyby = PopisChyby[1].Trim();
            Settings.Default.Save();
        }

        private void comboBox6_SelectionChangeCommitted(object sender, EventArgs e)
        {
            string[] Vinnik = comboBox6.SelectedItem.ToString().Split('•');  //uloží vinníka na odpis
            Settings.Default.vinnik = Vinnik[0].Trim();
            Settings.Default.Save();
        }

        private void checkBox1_CheckStateChanged(object sender, EventArgs e)
        {
            comboBox1.Items.Clear();
            comboBox2.Items.Clear();
            comboBox6.Items.Clear();
            label2.Visible = false;
            comboBox2.Visible = false;
            NacitajKomponent();
            comboBox6.Items.Clear();
            comboBox5.SelectedIndex = -1;
            textBox3.Text = 1.ToString();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)  //vyhladávač v textovom pre cbx - kód chyby
        {
             object FindItemContaining(IEnumerable items, string target)
             {
                foreach (object item in items)
                    if (item.ToString().Contains(target))
                        return item;
                return null;
             }
            comboBox5.SelectedItem = FindItemContaining(comboBox5.Items, textBox2.Text);
        }

        private void textBox2_Click(object sender, EventArgs e)
        {
            textBox2.SelectAll();
        }
    }
}
