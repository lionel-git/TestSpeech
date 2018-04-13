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


namespace TestSpeech
{
    public partial class Form1 : Form
    {
        
        SpeechSynthesizer _synth;

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

            //  _synth.SelectVoiceByHints(VoiceGender.Male);
            //   _synth.SelectVoice("Microsoft Paul");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var text=richTextBox1.SelectedText;
            if (string.IsNullOrEmpty(text))
                text = richTextBox1.Text;

            _synth.Rate = trackBar1.Value;
            var p = _synth.SpeakAsync(text);
        }

        private void LoadFile(string path)
        {
            richTextBox1.Text = File.ReadAllText(path, Encoding.Default);
            textBox1.Text = path;
        }


        private void button2_Click(object sender, EventArgs e)
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

        private void button3_Click(object sender, EventArgs e)
        {
            _synth.SpeakAsyncCancelAll();
        }
    }
}
