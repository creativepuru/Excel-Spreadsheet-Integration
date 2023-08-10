using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Excel_Spreadsheet_Integration
{
    public partial class formLogin : Form
    {
        public formLogin()
        {
            InitializeComponent();
        }

        private void Login_Load(object sender, EventArgs e)
        {

        }

        private void exitButton_Click(object sender, EventArgs e)
        {
            //Close();
            Application.Exit();
        }

        private void formLogin_Paint(object sender, PaintEventArgs e)
        {
            Graphics graphics = e.Graphics;

            Pen pen = new Pen(Color.Black);

            // Top side
            graphics.DrawLine(pen, 0, 0, Width, 0);
            // Bottom side
            graphics.DrawLine(pen, 0, Height - 1, Width, Height - 1);
            // Left side
            graphics.DrawLine(pen, 0, 0, 0, Height);
            // Right side
            graphics.DrawLine(pen, Width - 1, 0, Width - 1, Height);

            pen.Dispose();
        }
    }
}
