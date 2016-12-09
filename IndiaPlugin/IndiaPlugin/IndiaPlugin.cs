using System.Speech.Synthesis;
using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using KeePass.UI;
using KeePass;


using KeePass.Plugins;
using KeePass.Forms;
using KeePass.Resources;

using KeePassLib;
using KeePassLib.Security;

namespace IndiaPlugin
{
    public sealed class IndiaPluginExt : Plugin
    {
        private IPluginHost m_host = null;

        SpeechSynthesizer synthesizer = new SpeechSynthesizer();
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);
        
        // On creation of plugin, set-up resources.
        public override bool Initialize(IPluginHost host)
        {
            Terminate();
            if (host == null) return false;
            m_host = host;

            m_host.MainWindow.UIStateUpdated += this.OnUIStateUpdated;

            synthesizer.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Adult);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10
            synthesizer.SpeakAsync("Hi India. Type your password and press enter.");
            

            m_host.MainWindow.KeyPreview = true;
            m_host.MainWindow.KeyDown += SpeakSelectedEntry;
            m_host.MainWindow.KeyDown += SpeakAllEntries;
            m_host.MainWindow.KeyDown += HelpMenu;
            m_host.MainWindow.KeyDown += KillSpeech;

            m_host.MainWindow.EntryContextMenu.KeyDown += SpeakAddEntryMenu;
            m_host.MainWindow.EntryContextMenu.KeyDown += KillSpeech;


            return true;
       } 

        // Destruction of plugin, clean-up resources.
        public override void Terminate()
        {
            if (m_host == null) return;

            m_host.MainWindow.UIStateUpdated -= this.OnUIStateUpdated;

            m_host = null;
        }

        private void SpeakAddEntryMenu(object sender, KeyEventArgs e)
        {
            synthesizer.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Adult);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10

            ToolStripItemCollection tsMenu = m_host.MainWindow.EntryContextMenu.Items;

            if (e.KeyCode == Keys.I && e.Control && e.Shift)
            {
                synthesizer.SpeakAsync("yo");
                for (int i = 0; i < tsMenu.Count; i++)
                {
                    synthesizer.SpeakAsync(tsMenu[i].Text);
                }
                synthesizer.SpeakAsync(tsMenu.Count.ToString());
            }
        }

        private void HelpMenu(object sender, KeyEventArgs e)
        {
            synthesizer.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Adult);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10

            if (e.KeyCode == Keys.H && e.Control)
            {
                synthesizer.SpeakAsync("These are the keyboard options: press up or down to highlight an entry and have it spoken out, press control E to hear all the entries in the database, press control I to add a new entry, press control U to open the URL of the selected entry,press control H to hear this help menu again.");
            }
        }

        private void KillSpeech(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.K && e.Control && e.Shift) || e.KeyCode == Keys.Enter)
            {
                synthesizer.SpeakAsyncCancelAll();
            }
        }
        
        private void SpeakSelectedEntry(object sender, KeyEventArgs e)
        {
            synthesizer.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Adult);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10
            if (!m_host.Database.IsOpen)
            {
                synthesizer.SpeakAsync("You first need to open a database!");
                return;
            }

            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);
            System.Windows.Forms.ListView.SelectedIndexCollection selected_index = lv.SelectedIndices;

            if (e.KeyCode == Keys.Up)
            {
                synthesizer.SpeakAsyncCancelAll();
                for (int i = 0; i < lv.Items.Count; i++)
                {
                    if (selected_index.Contains(i))
                    {
                        if (i - 1 < 0)
                        {
                            synthesizer.SpeakAsync(lv.SelectedItems[0].Text);
                        }
                        else
                        {
                            synthesizer.SpeakAsync(lv.Items[i - 1].Text);
                        }
                    }
                }
            }
            if (e.KeyCode == Keys.Down)
            {
                synthesizer.SpeakAsyncCancelAll();
                if (lv.Items.Count == 0)
                {
                    lv.Focus();
                    lv.Items[0].Selected = true;
                }
                for (int i = 0; i < lv.Items.Count; i++)
                {
                    if (selected_index.Contains(i))
                    {
                        if (i + 1 >= lv.Items.Count)
                        {
                            synthesizer.SpeakAsync(lv.SelectedItems[0].Text);
                        }
                        else
                        {
                            synthesizer.SpeakAsync(lv.Items[i + 1].Text);
                        }
                    }
                }
            }
        }
        
        // Dictates all the items in the password listview.
        private void SpeakAllEntries(object sender, KeyEventArgs e)
        {
             
            synthesizer.SelectVoiceByHints(VoiceGender.Female, VoiceAge.Adult);
            synthesizer.Volume = 100;  // 0...100
            synthesizer.Rate = 0;     // -10...10
            
            if (!m_host.Database.IsOpen)
            {
                synthesizer.SpeakAsync("You first need to open a database!");
                return;
            }

            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);

            if (e.KeyCode == Keys.E && e.Control)
            {
                synthesizer.SpeakAsync("I am about to tell you the current entries.");
                for (int i = 0; i < lv.Items.Count; i++)
                {
                    synthesizer.SpeakAsync(lv.Items[i].Text);
                }
            }           

        }
        
        private void OpenURL(object sender, KeyEventArgs e)
        {
            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);

            if (e.KeyCode == Keys.U && e.Control)
            {
                //Open url
                System.Diagnostics.Process.Start(lv.SelectedItems[3].Text);
            }
        }
        private void SetSelectedEntryType(object sender, EventArgs e)
        // Dictates the tag for the currently selected entry in the listview.
        // TODO: Does not work as expected.
        {
            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);

            int index = lv.FocusedItem.Index;
        }

        // UI Refresh callback function.
        private void OnUIStateUpdated(object sender, EventArgs e)
        {
            ListView lv = (m_host.MainWindow.Controls.Find(
                "m_lvEntries", true)[0] as ListView);

            lv.BeginUpdate();
            /*
            for (int i = 0; i < lv.Items.Count; i++)
            {
                lv.Items[i].Font = new Font("Arial", 36);
            }*/
            for (int i = 0; i < lv.Items.Count; i++)
            {
                if (i % 2 == 0)
                {
                    lv.Items[i].BackColor = System.Drawing.Color.DarkSalmon;
                }
                else if (i % 2 == 1)
                {
                    lv.Items[i].BackColor = System.Drawing.Color.Blue;
                    lv.Items[i].ForeColor = System.Drawing.Color.WhiteSmoke;
                }
            }

            lv.EndUpdate();
        }
        
    }
}
