using Microsoft.Office.Interop.Excel;
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
    public partial class VyberProjektu : Form
    {
        public VyberProjektu()
        {
            InitializeComponent();
            Text = "Výber projektu na oddelení  " + Settings.Default.AktualneOddelenie;
            label1.Text = "Užívateľ:  " + Settings.Default.Aktualnyuzivatel;

            try   //načíta projekty do pomocného cbx
            {
                using (SqlConnection sqlConnection = new SqlConnection(Connection.ConnectionString))
                {
                    SqlCommand sqlCmd = new SqlCommand("SELECT Projekt FROM SM30 WHERE Oddelenie = '" + Settings.Default.AktualneOddelenie + "' AND Aktiviny1Vymazany0 = '1'", sqlConnection);
                    sqlConnection.Open();
                    SqlDataReader sqlReader = sqlCmd.ExecuteReader();
                    while (sqlReader.Read())
                    {
                        comboBox1.Items.Add(sqlReader["Projekt"].ToString());
                    }
                    sqlReader.Close();
                }
            }
            catch (Exception)
            { }

            object[] distinctItems = (from object o in comboBox1.Items select o).Distinct().ToArray();
            comboBox1.Items.Clear();
            comboBox1.Items.AddRange(distinctItems);

            int ii = 0;
            while (comboBox1.Items.Count - 1 >= ii)
            {
                if (Convert.ToString(comboBox1.Items[ii]).Trim() == string.Empty)
                {
                    comboBox1.Items.RemoveAt(ii);
                    ii -= 1;
                }
                ii += 1;
            }

            List<System.Windows.Forms.Button> buttonsToName = new List<System.Windows.Forms.Button>()
            {       //okrem button2 - tlacidlo spat
             button1,button3,button4,button5,button6,button7,button8,button9,button10,button11,button12,button13,button14,button15,button16,button17,button18,button19,
             button20,button21,button22,button23,button24,button25,button26,button27,button28,button29,button30,button31,button32,button33,button34,button35,button36,
             button37,button38,button39,button40,button41,button42,button43,button44,button45,button46,button47,button48,button49,button50,button51,button52,button53,
             button54,button55,button56,button57,button58,button59,button60,button61,button62,button63,button64,button65,button66,button67,button68,button69,button70,
             button71
            };  

            try
            {
                for (int i = 0; i < comboBox1.Items.Count;)
                {
                    foreach (var button in buttonsToName)
                    {
                        string value = comboBox1.GetItemText(comboBox1.Items[i]);
                        i += 1;
                        if (i > 70)
                        {break;}
                        button.Text = value;
                        button.Visible = true;
                    }
                }
            }
            catch (Exception){}
        }

        private void button2_Click(object sender, EventArgs e)  //Späť
        {
            Uvod nextForm = new Uvod();
            Hide();
            nextForm.ShowDialog();
            Close();
        }

        void NextPage()
        {
            _1 nextForm = new _1();
            Hide();
            nextForm.ShowDialog();
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button1.Text;Settings.Default.Save();NextPage();}
        private void button3_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button3.Text; Settings.Default.Save(); NextPage(); }
        private void button4_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button4.Text; Settings.Default.Save(); NextPage(); }
        private void button5_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button5.Text; Settings.Default.Save(); NextPage(); }
        private void button6_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button6.Text; Settings.Default.Save(); NextPage(); }
        private void button7_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button7.Text; Settings.Default.Save(); NextPage(); }
        private void button8_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button8.Text; Settings.Default.Save(); NextPage(); }
        private void button9_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button9.Text; Settings.Default.Save(); NextPage(); }
        private void button10_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button10.Text; Settings.Default.Save(); NextPage(); }
        private void button11_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button11.Text; Settings.Default.Save(); NextPage(); }
        private void button12_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button12.Text; Settings.Default.Save(); NextPage(); }
        private void button13_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button13.Text; Settings.Default.Save(); NextPage(); }
        private void button14_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button14.Text; Settings.Default.Save(); NextPage(); }
        private void button15_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button15.Text; Settings.Default.Save(); NextPage(); }
        private void button16_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button16.Text; Settings.Default.Save(); NextPage(); }
        private void button17_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button17.Text; Settings.Default.Save(); NextPage(); }
        private void button18_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button18.Text; Settings.Default.Save(); NextPage(); }
        private void button19_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button19.Text; Settings.Default.Save(); NextPage(); }
        private void button20_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button20.Text; Settings.Default.Save(); NextPage(); }
        private void button21_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button21.Text; Settings.Default.Save(); NextPage(); }
        private void button22_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button22.Text; Settings.Default.Save(); NextPage(); }
        private void button23_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button23.Text; Settings.Default.Save(); NextPage(); }
        private void button24_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button24.Text; Settings.Default.Save(); NextPage(); }
        private void button25_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button25.Text; Settings.Default.Save(); NextPage(); }
        private void button26_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button26.Text; Settings.Default.Save(); NextPage(); }
        private void button27_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button27.Text; Settings.Default.Save(); NextPage(); }
        private void button28_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button28.Text; Settings.Default.Save(); NextPage(); }
        private void button29_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button29.Text; Settings.Default.Save(); NextPage(); }
        private void button30_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button30.Text; Settings.Default.Save(); NextPage(); }
        private void button31_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button31.Text; Settings.Default.Save(); NextPage(); }
        private void button32_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button32.Text; Settings.Default.Save(); NextPage(); }
        private void button33_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button33.Text; Settings.Default.Save(); NextPage(); }
        private void button34_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button34.Text; Settings.Default.Save(); NextPage(); }
        private void button35_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button35.Text; Settings.Default.Save(); NextPage(); }
        private void button36_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button36.Text; Settings.Default.Save(); NextPage(); }
        private void button37_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button37.Text; Settings.Default.Save(); NextPage(); }
        private void button38_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button38.Text; Settings.Default.Save(); NextPage(); }
        private void button39_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button39.Text; Settings.Default.Save(); NextPage(); }
        private void button40_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button40.Text; Settings.Default.Save(); NextPage(); }
        private void button41_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button41.Text; Settings.Default.Save(); NextPage(); }
        private void button42_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button42.Text; Settings.Default.Save(); NextPage(); }
        private void button43_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button43.Text; Settings.Default.Save(); NextPage(); }
        private void button44_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button44.Text; Settings.Default.Save(); NextPage(); }
        private void button45_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button45.Text; Settings.Default.Save(); NextPage(); }
        private void button46_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button46.Text; Settings.Default.Save(); NextPage(); }
        private void button47_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button47.Text; Settings.Default.Save(); NextPage(); }
        private void button48_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button48.Text; Settings.Default.Save(); NextPage(); }
        private void button49_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button49.Text; Settings.Default.Save(); NextPage(); }
        private void button50_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button50.Text; Settings.Default.Save(); NextPage(); }
        private void button51_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button51.Text; Settings.Default.Save(); NextPage(); }
        private void button52_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button52.Text; Settings.Default.Save(); NextPage(); }
        private void button53_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button53.Text; Settings.Default.Save(); NextPage(); }
        private void button54_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button54.Text; Settings.Default.Save(); NextPage(); }
        private void button55_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button55.Text; Settings.Default.Save(); NextPage(); }
        private void button56_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button56.Text; Settings.Default.Save(); NextPage(); }
        private void button57_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button57.Text; Settings.Default.Save(); NextPage(); }
        private void button58_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button58.Text; Settings.Default.Save(); NextPage(); }
        private void button59_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button59.Text; Settings.Default.Save(); NextPage(); }
        private void button60_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button60.Text; Settings.Default.Save(); NextPage(); }
        private void button61_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button61.Text; Settings.Default.Save(); NextPage(); }
        private void button62_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button62.Text; Settings.Default.Save(); NextPage(); }
        private void button63_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button63.Text; Settings.Default.Save(); NextPage(); }
        private void button64_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button64.Text; Settings.Default.Save(); NextPage(); }
        private void button65_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button65.Text; Settings.Default.Save(); NextPage(); }
        private void button66_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button66.Text; Settings.Default.Save(); NextPage(); }
        private void button67_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button67.Text; Settings.Default.Save(); NextPage(); }
        private void button68_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button68.Text; Settings.Default.Save(); NextPage(); }
        private void button69_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button69.Text; Settings.Default.Save(); NextPage(); }
        private void button70_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button70.Text; Settings.Default.Save(); NextPage(); }
        private void button71_Click(object sender, EventArgs e)
        { Settings.Default.Aktualnyprojekt = button71.Text; Settings.Default.Save(); NextPage(); }
    }
}
