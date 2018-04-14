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

namespace TestSpeech
{
    public partial class Form1 : Form
    {
        
        private SpeechSynthesizer _synth;
        private Prompt _prompt;

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
            
            //  _synth.SelectVoiceByHints(VoiceGender.Male);
            //   _synth.SelectVoice("Microsoft Paul");
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


            _synth.Rate = trackBar1.Value;
            _synth.SelectVoice(comboBoxVoices.Text.Split('#')[0]);
            _prompt = _synth.SpeakAsync(text);
        }

        private void LoadFile(string path)
        {
            textBoxFileName.Text = path;
            if (Path.GetExtension(path).StartsWith(".doc"))
                LoadWordFile(path);
            else
                richTextBox1.Text = File.ReadAllText(path, Encoding.Default);
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
           
            var app = new Microsoft.Office.Interop.Word.Application();
            Document doc = app.Documents.Open(path);
            string words = doc.Content.Text;
            doc.Close();
            app.Quit();
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
    }
}
