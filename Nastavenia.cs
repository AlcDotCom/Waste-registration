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
    public partial class Nastavenia : Form
    {
        public Nastavenia()
        {
            InitializeComponent();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Dispose();
            Uvod nextForm = new Uvod();
            Hide();
            nextForm.ShowDialog();
            Close();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            Dispose();
            Nastavenia_VstupneData nextForm = new Nastavenia_VstupneData();
            Hide();
            nextForm.ShowDialog();
            Close();
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            Dispose();
            Adresa nextForm = new Adresa();
            Hide();
            nextForm.ShowDialog();
            Close();
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            Dispose();
            KodChyby nextForm = new KodChyby();
            Hide();
            nextForm.ShowDialog();
            Close();
        }
        public new void Dispose()
        {
            pictureBox2.Image.Dispose();
            pictureBox2.Image = null;
            pictureBox3.Image.Dispose();
            pictureBox3.Image = null;
            pictureBox4.Image.Dispose();
            pictureBox4.Image = null;
            pictureBox5.Image.Dispose();
            pictureBox5.Image = null;
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            Dispose();
            Nastavenia__UpravaDat nextForm = new Nastavenia__UpravaDat();
            Hide();
            nextForm.ShowDialog();
            Close();
        }
    }
}
