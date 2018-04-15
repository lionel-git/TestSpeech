using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Speech.Synthesis;
using System.IO;
using Microsoft.Office.Interop.Word;
using System.Reflection;

namespace TestSpeech
{
    public partial class Form1 : Form
    {
        
        private SpeechSynthesizer _synth;
        private Prompt _prompt;

        private Microsoft.Office.Interop.Word.Application _wordApp;

        private List<string> _voices;

        public Form1()
        {
            InitializeComponent();

            // Initialize a new instance of the SpeechSynthesizer.
            _synth = new SpeechSynthesizer();
            // Configure the audio output. 
            _synth.SetOutputToDefaultAudioDevice();


            var l=_synth.GetInstalledVoices();

            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);

            _voices = new List<string>();
            _voices.AddRange(l.Select(x => 
            $"{x.VoiceInfo.Name}# {x.VoiceInfo.Gender} {x.VoiceInfo.Culture}"));
            comboBoxVoices.DataSource = _voices;

            var wordVersion = IsWordInteropInstalled();
            if (wordVersion != null)
                toolStripStatusLabel1.Text = $"Word interpop found: {wordVersion}";
            else
                toolStripStatusLabel1.Text = "Warning: Word interop not found";

            Text = $"Test Speech {Assembly.GetExecutingAssembly().GetName().Version.ToString()}";

           // notifyIcon1.ContextMenu = contextMenuStrip1;
        }


        private void buttonRead_Click(object sender, EventArgs e)
        {
            var text=richTextBox1.SelectedText;
            if (string.IsNullOrEmpty(text))
            {
                text = richTextBox1.Text;
                if (checkBoxFromCursor.Checked)
                    text = text.Substring(richTextBox1.SelectionStart, text.Length - richTextBox1.SelectionStart);
            }

            if (string.IsNullOrEmpty(text))
            {
                MessageBox.Show("No text selected", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                _synth.Rate = trackBar1.Value;
                _synth.SelectVoice(comboBoxVoices.Text.Split('#')[0]);
                _prompt = _synth.SpeakAsync(text);
            }
        }

        private void LoadFile(string path)
        {
            textBoxFileName.Text = path;
            var extension = Path.GetExtension(path);
            Cursor.Current = Cursors.WaitCursor;
            if (extension.StartsWith(".doc") || extension==".odt")
                LoadWordFile(path);
            else
                richTextBox1.Text = File.ReadAllText(path, Encoding.Default);
            Cursor.Current = Cursors.Default;
        }


        private void buttonOpenFile_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
                LoadFile(openFileDialog1.FileName);
        }

        void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
                LoadFile(file);
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            _synth.SpeakAsyncCancelAll();
        }

        private void LoadWordFile(string path)
        {
            if (_wordApp == null)
            {
                _wordApp = new Microsoft.Office.Interop.Word.Application();
                toolStripStatusLabel2.Text = $"Word version: {_wordApp.Build}";
            }
            Document doc = _wordApp.Documents.Open(path);
            string words = doc.Content.Text;
            doc.Close();
            richTextBox1.Text = words;
        }

        private void buttonPauseResume_Click(object sender, EventArgs e)
        {
            switch (_synth.State)
            {
                case SynthesizerState.Paused:
                    _synth.Resume();
                    break;
                case SynthesizerState.Speaking:
                    _synth.Pause();
                    break;
            }
        }

        private static string IsWordInteropInstalled()
        {
            Type officeType = Type.GetTypeFromProgID("Word.Application");
            if (officeType != null)
            {
                return $"{officeType.Assembly.FullName}";
            }
            else
                return null;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (_wordApp != null)
                _wordApp.Quit();
        }
    }
}
