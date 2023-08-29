using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Waste_registration
{
    public partial class heslo : Form
    {
        public heslo()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Uvod nextForm = new Uvod();
            Hide();
            nextForm.ShowDialog();
            Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "HeslO")
            {
                Nastavenia nextForm = new Nastavenia();
                Hide();
                nextForm.ShowDialog();
                Close();
            }
            else
            {
                MessageBox.Show("Nesprávne heslo");
                textBox1.Text = "";
                textBox1.Focus();
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button1.PerformClick();
            }
        }
    }
}
