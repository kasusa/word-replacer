using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using archiver.ConsoleColorWriter;

namespace archiver
{
    public partial class Form_replaceForAll : Form
    {
        public string path_stub;
        public Form_replaceForAll()
        {
            InitializeComponent(); AllocConsole();
            Console.WriteLine(@"
Dont close cmd window unless you want to exit.
Only docx files are supported

Help:
Drag/paste the directory or docxfile into 【floder path】or【Docs】. Then all docx files in the directory can be automatically recognized
Use shift/ctrl to multi-select the Docs , remove them using buttons on the right. 
file will not be delete but only remove from box.
all files in the Docs box will be operated. 
click Replace to replace (you can choose save location)

CMD mode is more powerful and can replace muti pair of words in same time.
you can also watch a video tutroal here:
");
        }
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();  
        #region textDrop
        private void textBox1_DragDrop(object sender, DragEventArgs e)
        {
            textBox1.Text = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            if (textBox1.Text.EndsWith(".docx"))//如果你把文件夹中的.docx文件拖进来，也可以自动识别文件夹目录
            {
                var mystr = textBox1.Text;
                var laststr = mystr.Split('\\')[mystr.Split('\\').Length - 1];
                textBox1.Text = mystr.Replace("\\" + laststr, ""); ;
            }
            string path = textBox1.Text;

        }

        private void textBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Link;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            button2.PerformClick();
        }
        #endregion



        //一键替换
        private void button1_Click(object sender, EventArgs e)
        {

            if (textBox2 .Text == "")
            {
                toolStripStatusLabel1.Text = "I cant search for nothing.";
            }
            this.Hide();
            ConsoleWriter.WriteCyan(textBox2.Text + "→" + textBox3.Text);

            foreach (var item in listBox1.Items)
            {
                string filenam = item.ToString();
                myutil doc = new myutil(path_stub + filenam);
                doc._replacePatterns.Add(textBox2.Text , textBox3 .Text);
                doc.ReplaceTextWithText_all_noBracket();
                Console.WriteLine("已经处理："+filenam);
                if (radioButton1.Checked)
                {
                    doc.save();
                    toolStripStatusLabel1.Text = "saved to desktop\\out";
                }
                else
                {
                    doc.saveUrl(path_stub+filenam);
                    toolStripStatusLabel1.Text = "saved , replaced source files";

                }
            }
            Console.WriteLine("done.");
            ConsoleWriter.WriteSeperator('#');
            this.Show();
        }

        //命令行模式
        private void button3_Click(object sender, EventArgs e)
        {
            if (true)
            {

                if (textBox1.Text == "")
                {
                    toolStripStatusLabel1.Text = "select file first!~。";
                    return;
                }
                this.Hide();
                Dictionary<string, string> map = new Dictionary<string, string>();
                ConsoleWriter.WriteSeperator('-');
                ConsoleWriter.WriteColoredText("Please enter the replacement content, multiple lines are supported, separated by | (vertical bar):", ConsoleColor.Green);
                while (true)
                {
                    var line = Console.ReadLine().Split("|");
                    if (line[0] == "") break;
                    if (line.Length == 1)
                    {
                        ConsoleWriter.WriteYEllow("No vertical bar detected, please re-enter correctly");
                        continue;
                    }
                    map.Add(line[0], line[1]);
                    ConsoleWriter.WriteCyan("Successfully write to dictionary, continue to add more (enter to leave)");
                }

                Console.WriteLine("Get the dictionary:");

                foreach (var kvp in map)
                {
                    ConsoleWriter.WriteCyan(kvp.Key + " → " + kvp.Value);
                }
                Console.WriteLine("will replace:" + map.Count);


                foreach (var item in listBox1.Items)
                {
                    string filenam = item.ToString();
                    myutil doc = new myutil(path_stub + filenam);
                    doc._replacePatterns = map;

                    doc.ReplaceTextWithText_all_noBracket();
                    Console.WriteLine("processing：" + filenam);
                    if (radioButton1.Checked)
                    {
                        doc.save();
                        toolStripStatusLabel1.Text = "saved to desktop\\out";
                    }
                    else
                    {
                        doc.saveUrl(path_stub + filenam);
                        toolStripStatusLabel1.Text = "saved , replaced source files";

                    }
                }
                ConsoleWriter.WriteSeperator('#');
                this.Show();
            }
           
        }

        //刷新列表
        private void button2_Click(object sender, EventArgs e)
        {
            string path = textBox1.Text;
            if (!Directory.Exists(path))
            {
                Console.WriteLine("dir is not correct");
                return;
            }
            else
            {
                ConsoleWriter.WriteColoredText("dir is detected：" + textBox1.Text,ConsoleColor.Green);

                path_stub = path + "\\";
                string[] fangan_list = Directory.GetFiles(path, "*.docx");
                
                listBox1.Items.Clear();
                foreach (var item in fangan_list)
                {
                    if (!item.StartsWith("~"))
                    {
                        listBox1.Items.Add(item.Replace(path + "\\", ""));
                    }
                }
            }
            //string text = textBox1.Text;
            //string filename = text.Split('\\')[text.Split('\\').Length - 1];
        }

        #region savelocation

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                radioButton2.Checked = false;
            }
            else 
            {
                radioButton2.Checked = true;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                radioButton1.Checked = false;
            }
            else
            {
                radioButton1.Checked = true;
            }
        }
        #endregion


        //移除选中项目
        private void button4_Click(object sender, EventArgs e)
        {
            List<int> f = new List<int>();
            for (int i = listBox1.Items.Count - 1; i >=0; i--)
            {
                if (listBox1.SelectedIndices.Contains(i))
                {
                    listBox1.Items.RemoveAt(i);
                }
            }
            f.Reverse();
        }

        //仅保留选中项目
        private void button5_Click(object sender, EventArgs e)
        {
            List<int> f = new List<int>();
            for (int i = listBox1.Items.Count - 1; i >= 0; i--)
            {
                if (!listBox1.SelectedIndices.Contains(i))
                {
                    listBox1.Items.RemoveAt(i);
                }
            }
            f.Reverse();
        }

        //表格杠杠
        private void button6_Click(object sender, EventArgs e)
        {
            this.Hide();
            if (listBox1.Items.Count ==0)
            {
                toolStripStatusLabel1.Text = "巧妇难为无米之炊";
            }
            foreach (var item in listBox1.Items)
            {
                ConsoleWriter.WriteCyan(textBox2.Text + "→" + textBox3.Text);
                string filenam = item.ToString();
                myutil doc = new myutil(path_stub + filenam);

                Console.WriteLine("杠杠：" + filenam);
                doc.Table_gang();

                if (radioButton1.Checked)
                {
                    doc.save();
                    toolStripStatusLabel1.Text = "处理完毕(gang)，保存到桌面out";
                }
                else
                {
                    doc.saveUrl(path_stub + filenam);
                    toolStripStatusLabel1.Text = "处理完毕(gang)，替换掉了文件";

                }
            }
            ConsoleWriter.WriteGray("处理完毕");
            ConsoleWriter.WriteSeperator('#');
            this.Show();
        }

        //一个临时文字板
        private void toolStripSplitButton1_ButtonClick(object sender, EventArgs e)
        {
            if (File.Exists("changeto.txt"))
            {
                System.Diagnostics.Process.Start("notepad.exe", "changeto.txt");
            }
            else
            {
                var f = File.Create("changeto.txt");
                f.Close();
                System.Diagnostics.Process.Start("notepad.exe", "changeto.txt");
            }
            
        }

        private void button7_Click(object sender, EventArgs e)
        {
            textBox2.Text = textBox3.Text = "";
        }

        private void 命令行模式ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            button3.PerformClick();
        }
    }
}
