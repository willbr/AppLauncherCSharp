﻿using System;
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
        private bool Started = false;
        private bool CommandMode = false;
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
                this.Started = true;
                this.Visible = !this.Visible;
            }
            base.WndProc(ref m);
        }

        // http://stackoverflow.com/questions/5168249/c-showing-an-invisible-form
        // skips first show event
        protected override void SetVisibleCore(bool value)
        {
            if (!this.IsHandleCreated && !this.Started)
            {
                this.CreateHandle();
                value = false;   // Prevent window from becoming visible
            }
            if (!this.Started)
                value = false;
            base.SetVisibleCore(value);
        }

        public Form1()
        {
            InitializeComponent();
            bool r = RegisterHotKey(this.Handle, HOTKEY_ID, MOD.MOD_CONTROL, (int)' ');
            Debug.WriteLine(string.Format("reg hk {0}", r));
            if (r == false)
            {
                MessageBox.Show("failed to register hotkey");
                this.Close();
            }
        }

        private void Form1_Load(
            object sender,
            EventArgs e
            )
        {
            IndexLinks();
            FilterLinks();
        }

        private void textBox1_TextChanged(
            object sender,
            EventArgs e
            )
        {
            string text = textBox1.Text;
            if (text.Length > 0 && @";+-*/".IndexOf(text[0]) != -1)
            {
                CommandMode = true;
                ParseCommand();
            }
            else
            {
                CommandMode = false;
                FilterLinks();
            }
        }

        private void Form1_KeyDown(
            object sender,
            KeyEventArgs e
            )
        {
            switch (e.KeyCode)
            {
                case Keys.Escape:
                    Hide();
                    e.SuppressKeyPress = true; // stops beeping
                    break;
                case Keys.Enter:
                    if (CommandMode)
                    {
                        if (textBox1.Text[0] == ';')
                        {
                            string cmd = textBox1.Text.Substring(1);
                            switch (cmd)
                            {
                                case "q":
                                case "quad":
                                    Close();
                                    break;
                                default:
                                    Debug.WriteLine("unknown command: " + cmd);
                                    break;
                            }
                        }
                        else
                        {
                            Clipboard.SetText((string)listBox1.SelectedItem);
                        }
                    }
                    else
                    {
                        StartMenuLinks.Launch((string)listBox1.SelectedItem);
                    }
                    Hide();
                    e.SuppressKeyPress = true;
                    break;
                case Keys.Tab:
                    
                    if (e.Shift)
                    {
                        Debug.WriteLine("+tab");
                        PreviousSelection();
                    }
                    else
                    {
                        Debug.WriteLine("tab");
                        NextSelection();
                    }
                    e.SuppressKeyPress = true;
                    break;
                default:
                    textBox1.Focus();
                    break;
            }
        }

        private void Form1_Deactivate(object sender, EventArgs e)
        {
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

        private void clearToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

        private void homeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.SelectionStart = 0;
        }

        private void endToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.SelectionStart = textBox1.TextLength;
        }

        private void nextToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NextSelection();
        }

        private void previusToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PreviousSelection();
        }

        private void forwardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ForwardChar();
        }

        private void backwardToolStripMenuItem_Click(object sender, EventArgs e)
        {
            BackwardChar();
        }

        private void killToolStripMenuItem_Click(object sender, EventArgs e)
        {
            textBox1.SelectionLength = textBox1.TextLength;
            textBox1.SelectedText = "";
        }

        private void BackwardChar()
        {
            if (textBox1.SelectionStart > 0)
                textBox1.SelectionStart -= 1;
        }

        private void ForwardChar()
        {
            if (textBox1.SelectionStart < textBox1.TextLength)
                textBox1.SelectionStart += 1;
        }

        private void NextSelection()
        {
            if (listBox1.SelectedIndex < listBox1.Items.Count)
                listBox1.SelectedIndex += 1;
        }

        private void PreviousSelection()
        {
            if (listBox1.SelectedIndex > 0)
                listBox1.SelectedIndex -= 1;
        }

        private void Form1_VisibleChanged(object sender, EventArgs e)
        {
            textBox1.Focus();
            // Select all text
            textBox1.SelectionStart = 0;
            textBox1.SelectionLength = textBox1.TextLength;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            textBox1.Focus();
        }

        private void ParseCommand()
        {
            if (textBox1.Text[0] == ';')
            {
                string cmd = textBox1.Text.Substring(1);
            }
            else
            {
                Debug.WriteLine("parse math");
                string cmd = textBox1.Text;
                Stack<string> s = new Stack<string>();
                List<string> l = new List<string>();
                string[] tokens = Regex.Split(Regex.Replace(cmd, @"([\(\)])", @" $1 "), @"\s+");
                foreach (string token in tokens)
                {
                    try
                    {
                        if (token == ")")
                        {
                            bool foundPair = false;
                            while (foundPair == false)
                            {
                                string lt = s.Pop();

                                if (lt == "(")
                                {
                                    foundPair = true;
                                    s.Push(EvaluateList(l));
                                }
                                else
                                {
                                    l.Insert(0, lt);
                                }
                                Debug.WriteLine("l: " + String.Join(", ", l));
                            }
                        }
                        else
                        {
                            if (token != "") s.Push(token);
                        }
                    }
                    catch (InvalidOperationException e)
                    {
                        Debug.WriteLine("unbalanced math equations");
                    }
                }
                Debug.WriteLine("s: " + String.Join(", ", s));
                listBox1.Items.Clear();
                try
                {
                    string r = EvaluateList(s.Reverse().ToList());
                    Debug.WriteLine("r: " + r);

                    listBox1.Items.Add(r);
                }
                catch
                {
                    listBox1.Items.Add("error: math error");
                }
                finally
                {
                    listBox1.SelectedIndex = 0;
                    // TODO append math history
                }
            }
        }

        private string EvaluateList(List<string> list)
        {
            if (list.Count == 0) return "";
            float a = 0;
            string command = list[0];
            list.RemoveAt(0);
            switch (command)
            {
                case "+":
                    a = 0;
                    foreach (string token in list)
                        a += float.Parse(token);
                    break;
                case "-":
                    a = float.Parse(list[0]);
                    list.RemoveAt(0);
                    foreach (string token in list)
                        a -= float.Parse(token);
                    break;
                case "*":
                    a = 1;
                    foreach (string token in list)
                        a *= float.Parse(token);
                    break;
                case "/":
                    a = float.Parse(list[0]);
                    list.RemoveAt(0);
                    foreach (string token in list)
                        a /= float.Parse(token);
                    break;
                default:
                    break;
            }
            return a.ToString();
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
        private string[] IgnoreTerms = {
                                    "uninstall",
                                    "read ?me",
                                    "manual",
                                    "about",
                                    "help",
                                    "license",
                                    "reset",
                                };

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
                        Link link = new Link(filePath, 0);

                        if (ValidLink(link))
                            Links.Add(link);
                    }
                }
                catch (Exception e)
                {
                    Debug.WriteLine("AddLinks exception: " + e);
                }
            }
        }

        private bool ValidLink(Link link)
        {
            foreach(string term in IgnoreTerms)
            {
                if (Regex.IsMatch(link.FileName, term, RegexOptions.IgnoreCase))
                {
                    return false;
                }
            }
            return true;
        }

        public List<Link> Filter(string needle)
        {
            // TODO valid needle for regex
            string pattern = needle;
            // TODO score and sort
            List<Link> matches = new List<Link>();
            try
            {
                foreach (Link link in Links)
                {
                    if (Regex.IsMatch(link.FileName, pattern, RegexOptions.IgnoreCase))
                    {
                        matches.Add(link);
                    }
                }
            }
            catch (ArgumentException e)
            {
                Debug.WriteLine("error regex: " + e);
            }

            return matches;
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
