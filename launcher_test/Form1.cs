using System;
using System.Diagnostics;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace launcher_test
{
    public partial class Form1 : Form
    {
        private const int WM_HOTKEY = 0x0312;
        private const int HOTKEY_ID = 0;
        private WindowsLinks StartMenuLinks = new WindowsLinks();

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
            Debug.WriteLine("load");
            bool r = RegisterHotKey(this.Handle, HOTKEY_ID, MOD.MOD_CONTROL, (int)' ');
            Debug.WriteLine(string.Format("reg hk {0}", r));
            if (r == false)
            {
                MessageBox.Show("failed to register hotkey");
                this.Close();
            }
            IndexLinks();
            FilterLinks();
        }

        private void textBox1_TextChanged(
            object sender,
            EventArgs e
            )
        {
            FilterLinks();
        }

        private void Form1_KeyDown(
            object sender,
            KeyEventArgs e
            )
        {
            if (e.Control)
            {
                Debug.WriteLine("dn ctrl");

                switch (e.KeyCode)
                {
                    case Keys.A:
                        textBox1.SelectionStart = 0;
                        break;
                    case Keys.E:
                        textBox1.SelectionStart = textBox1.TextLength;
                        break;
                    case Keys.U:
                        textBox1.Text = "";
                        break;
                    case Keys.N:
                        if (listBox1.SelectedIndex < (listBox1.Items.Count - 1))
                        {
                            listBox1.SelectedIndex += 1;
                        }
                        break;
                    case Keys.P:
                        if (listBox1.SelectedIndex > 0)
                        {
                            listBox1.SelectedIndex -= 1;
                        }
                        break;
                    default:
                        break;
                }
                return;
            }
            else
            {
                switch (e.KeyCode)
                {
                    case Keys.Escape:
                        //Debug.WriteLine("escape");
                        //Close();
                        this.Hide();
                        break;
                    case Keys.Enter:
                        //Debug.WriteLine("enter");
                        //Debug.WriteLine((string)listBox1.SelectedItem);
                        StartMenuLinks.Launch((string)listBox1.SelectedItem);
                        Hide();
                        break;
                    default:
                        break;
                }
            }
        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {
            //Debug.WriteLine("deactivate");
            Hide();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID);
        }

        private void IndexLinks()
        {
            string commonStartMenu = Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu);
            string userStartMenu = Environment.GetFolderPath(Environment.SpecialFolder.StartMenu);
            StartMenuLinks.Folders.Add(commonStartMenu);
            StartMenuLinks.Folders.Add(userStartMenu);
            StartMenuLinks.Index();
        }

        private void FilterLinks()
        {
            listBox1.Items.Clear();
            foreach (Link link in StartMenuLinks.Filter(textBox1.Text))
            {
                listBox1.Items.Add(link.FileName);
            }
            if (listBox1.Items.Count > 0)
            {
                listBox1.SelectedIndex = 0;
            }
        }
    }

    public class Link
    {
        public string FileName;
        public string FilePath;
        public int Score;

        public Link(string filePath, int score)
        {
            FilePath = filePath;
            FileName = Path.GetFileNameWithoutExtension(filePath);
            Score = score;
        }
    }

    public class WindowsLinks
    {
        public List<string> Folders = new List<string>();
        private List<Link> Links;

        public void Index()
        {
            Links = new List<Link>();
            foreach (string folder in Folders)
            {
                try
                {
                    var lnkFiles = Directory.EnumerateFiles(folder, "*.lnk", SearchOption.AllDirectories);
                    foreach (string filePath in lnkFiles)
                    {
                        // TODO load history score
                        Links.Add(new Link(filePath, 0));
                    }
                }
                catch (Exception e)
                {
                    Debug.WriteLine("AddLinks exception: " + e);
                }
            }
        }

        public List<Link> Filter(string needle)
        {
            // TODO valid needle for regex
            string pattern = needle;
            // TODO score and sort
            List<Link> matches = new List<Link>();
            foreach(Link link in Links)
            {
                if (Regex.IsMatch(link.FileName, pattern, RegexOptions.IgnoreCase))
                {
                    matches.Add(link);
                }
            }

            if (matches.Count > 5)
            {
                return matches.GetRange(0, 5);
            }
            else
            {
                return matches;
            }
        }

        public void Launch(string fileName)
        {
            foreach(Link link in Links)
            {
                if (link.FileName == fileName)
                {
                    System.Diagnostics.Process.Start(link.FilePath);
                    break;
                }
            }
        }
    }

}
