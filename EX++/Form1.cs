using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.Json;
using Newtonsoft.Json;
using System.Threading;

namespace ExPlus
{
    public partial class Form1
        : Form
    {

        VoiceProcessing vp;
        Amplifier amplifier;
        Filter filter;

        public Form1()
        {
            InitializeComponent();
            initHandles();
            Macros.UnprotectWindow(hwn);

        }

        IntPtr hwn;
        IntPtr htree;
        void initHandles()
        {
            var hndls = BackEnd.HandleProc();
            hwn = hndls.Item1;
            htree = hndls.Item2;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            initHandles();

            var n = Macros.GetSelectedInstrumentNum(hwn, htree);
            if (n < 0) return;

            Macros.ProtectWindow(hwn);
            Macros.SelectAllVoices(hwn, htree, n);

            vp = Macros.GetVoiceProcessingDatas(hwn, htree);
            //filter = Macros.getFilter(hwn);
            
            Macros.UnprotectWindow(hwn);
            var json = JsonConvert.SerializeObject(vp, Formatting.Indented);
            MessageBox.Show(json);

            //amplifier = Macros.GetAmplifier(hwn);
            //var json2 = JsonConvert.SerializeObject(a, Formatting.Indented);
            //MessageBox.Show(json2);



        }

        
        private void button4_Click(object sender, EventArgs e)
        {
            var n = Macros.GetSelectedInstrumentNum(hwn, htree);
            if (n < 0) return;

            Macros.ProtectWindow(hwn);
            Macros.SetVoiceProcessingDatas(hwn, htree, ref vp);
            //Macros.SetFilter(hwn, filter);
            //Macros.SetAmplifier(hwn, amplifier);
            Macros.UnprotectWindow(hwn);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SelectedPreset.Text = Macros.CreateNewPreset(hwn, htree).ToString();
        }
        private void button6_Click(object sender, EventArgs e)
        {
          
        }

        private void button6_Click_1(object sender, EventArgs e)
        {

            var sp = Macros.GetSelectedInstrumentNum(hwn, htree);
            if (sp < 0) return;

            Macros.PresetLinksToNewPreset(hwn, htree, sp );
        }
    }
}
