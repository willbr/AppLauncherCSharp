using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace launcher_test
{
    public partial class Form1 : Form
    {
        private const int WM_HOTKEY = 0x0312;
        private const int HOTKEY_ID = 0;

        [Flags]
        private enum MOD : uint
        {
            MOD_ALT = 0x01,
            MOD_CONTROL = 0x02,
            MOD_SHIFT = 0x04,
            MOD_WIN = 0x08
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool RegisterHotKey(
            IntPtr hWnd,
            int id,
            MOD fsModifiers,
            uint vk);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        protected override void WndProc(
            ref Message m
            )
        {
            if (m.Msg == WM_HOTKEY) {
                int id = m.WParam.ToInt32();
                //MessageBox.Show(string.Format("Hotkey {0}", id));
                this.Show();
            }
            base.WndProc(ref m);
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(
            object sender,
            EventArgs e
            )
        {
            this.Visible = true;
            Log("load");
            bool r = RegisterHotKey(this.Handle, HOTKEY_ID, MOD.MOD_CONTROL, (int)' ');
            Log(string.Format("reg hk {0}", r));
            if (r == false)
            {
                MessageBox.Show("failed to register hotkey");
                this.Close();
            }
            listBox1.Items.Add("hello");
            listBox1.Items.Add("two");
            listBox1.Items.Add("see");

            IndexLinks();
        }

        private void Log(
            string message
            )
        {
            textBox2.Text += message + Environment.NewLine;
            textBox2.SelectionStart = textBox2.Text.Length;
            textBox2.ScrollToCaret();
        }

        private void textBox1_TextChanged(
            object sender,
            EventArgs e
            )
        {
            Log("change");
        }

        private void Form1_KeyUp(
            object sender,
            KeyEventArgs e
            )
        {
            Log("up");

            switch (e.KeyCode)
            {
                case Keys.Escape:
                    Log("escape");
                    Close();
                    //this.Hide();
                    break;
                case Keys.Enter:
                    Log("enter");
                    Log((string)listBox1.SelectedItem);
                    listBox1.Items.Clear();
                    break;
                case Keys.Space:
                    Log("space");
                    break;
                default:
                    break;
            }

            if (e.Control)
            {
                Log("control");
            }
            else
            {
                ;
            }
        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {
            Log("deactivate");
            this.Hide();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            //MessageBox.Show("closed");
            UnregisterHotKey(this.Handle, HOTKEY_ID);
        }

        private void IndexLinks()
        {
            string CommonStartMenu = Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu);
            string UserStartMenu = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu);
            AddLinks(CommonStartMenu);
            AddLinks(UserStartMenu);
        }

        private void AddLinks(string folder)
        {
            try
            {
                var lnkFiles = Directory.EnumerateFiles(folder, "*.lnk", SearchOption.AllDirectories);

                foreach (string currentFile in lnkFiles)
                {
                    Log(Path.GetFileNameWithoutExtension(currentFile));
                }
            }
            catch (Exception e)
            {
                Log("AddLinks exception: " + e);
            }
        }
    }
}
