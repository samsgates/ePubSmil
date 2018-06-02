using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.IO;
using WaveLib;
using Yeti.MMedia.Mp3;
using System.Text.RegularExpressions;
using System.Speech.Synthesis;
using System.Globalization;

namespace ePubSmil
{
    public partial class MForm : ComponentFactory.Krypton.Toolkit.KryptonForm
    {
       
        string curErrLog = "";
      
       
        BackgroundWorker _bw;
        delegate void update_probar(string text, int max, int cur);
        bool audioComplete = false;

        string spanTxt = "";
        string smilTxt = "";
        string last_time = "";
        int spanID = 1;
        string curPageHtml = "";
        string curaudioFile = "";
        string curFullTxt = "";
        string htLineTxt = "";
        int spanPoint = 0;
        string voiceName = "";
        ArrayList cpList = new ArrayList();
        ArrayList attList = new ArrayList();

        public MForm()
        {
            InitializeComponent();
        }

        private void MForm_Load(object sender, EventArgs e)
        {
            this.Text += " " + Application.ProductVersion;

            load_voice_list();
        }

        public void load_voice_list()
        {

            cmb_voicelist.Items.Clear();
            using (SpeechSynthesizer synthesizer = new SpeechSynthesizer())
            {
                // Output information about all of the installed voices that
                // support the en-US locacale. 

                foreach (InstalledVoice voice in synthesizer.GetInstalledVoices(new CultureInfo("en-US")))
                {
                    VoiceInfo info = voice.VoiceInfo;
                    cmb_voicelist.Items.Add(info.Name);
                }

            }
            if (cmb_voicelist.Items.Count > 0)
            {
                cmb_voicelist.SelectedIndex = 0;
            }

        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            if (txt_input.Text == "") {
                gCls.show_error("Paste input content");
                return;
            }
            if (txt_savefile.Text == "") {
                gCls.show_error("Select MP3 file save path");
                return;
            }


            convert_smil();
        }

        public void convert_smil() {

            try
            {
                string tempDir = Application.StartupPath + "\\audio";
                if (Directory.Exists(tempDir)) {
                    Directory.Delete(tempDir, true);
                }
                Directory.CreateDirectory(tempDir);


                htLineTxt = txt_input.Text;

                htLineTxt = htLineTxt.Replace("\r\n", " ");
                htLineTxt = htLineTxt.Replace("\r", " ");
                htLineTxt = htLineTxt.Replace("\n", " ");
                htLineTxt = Regex.Replace(htLineTxt, "\\s+", " ");
                
                string readLine = "";
                readLine = htLineTxt;
                readLine = StripTagsRegex(readLine);
                readLine = Regex.Replace(readLine, "\\s+", " ");

                read_aloud_html(readLine, tempDir, cmb_voicelist.Text, (int) txt_sid.Value );

            }
            catch (Exception erd) {
                gCls.show_error(erd.Message.ToString());
                return;
            }
        }
        public string StripTagsRegex(string source)
        {
            return Regex.Replace(source, "<.*?>", string.Empty);
        }

        public void read_aloud_html(string inHtml, string audir,string voicename,int startID)
        {
            try
            {


                System.Speech.Synthesis.SpeechSynthesizer l_spv = new System.Speech.Synthesis.SpeechSynthesizer();
                l_spv.SpeakCompleted += new EventHandler<System.Speech.Synthesis.SpeakCompletedEventArgs>(spv_SpeakCompleted);
                l_spv.SpeakProgress += new EventHandler<System.Speech.Synthesis.SpeakProgressEventArgs>(spv_SpeakProgress);

                l_spv.SelectVoice(voicename);

                l_spv.Rate = -2;
                l_spv.SetOutputToWaveFile(audir + "\\audio.wav");
                audioComplete = false;
                spanTxt = "";
                smilTxt = "";
                last_time = "";
                curFullTxt = "";
                              

                curFullTxt = inHtml;
                spanID = startID;
                spanPoint = 0;
                curPageHtml = "page.html";
                curaudioFile = "audio.mp3";

                probar.Visible = true;
                probar.Maximum = Regex.Matches(inHtml, "\\s").Count;
                probar.Value = 0;



                l_spv.SpeakAsync(inHtml);
                while (audioComplete == false)
                {
                    Application.DoEvents();
                }

               

                string wSmil = "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n";
                wSmil += "<smil xmlns=\"http://www.w3.org/ns/SMIL\" version=\"3.0\" profile=\"http://www.idpf.org/epub/30/profile/content/\">\n";
                wSmil += "<body>\n";
                wSmil += smilTxt;
                wSmil += "\n</body></smil>\n";
                txt_smil.Text = wSmil;
                txt_out.Text = htLineTxt;
               // File.WriteAllText(htmlDir + "\\page" + pgNum + ".smil", wSmil);

                probar.Visible = false;
                l_spv.Dispose();

                //write mp3 file
                convert_mp3(audir);
                if (File.Exists(txt_savefile.Text)) {
                    File.Delete(txt_savefile.Text);
                }
                File.Copy(audir + "\\audio.mp3", txt_savefile.Text, true);

                gCls.show_message("SMIL Tag generated successfully");

            }
            catch(Exception erd)
            {
                gCls.show_error(erd.Message.ToString());
                return;
            }


        }

        private void spv_SpeakCompleted(object sender, System.Speech.Synthesis.SpeakCompletedEventArgs e)
        {
            audioComplete = true;
            string outsmile = smilTxt;
            outsmile = outsmile.Replace("[etime/]", last_time);
            smilTxt = outsmile;

        }
        private void spv_SpeakProgress(object sender, System.Speech.Synthesis.SpeakProgressEventArgs e)
        {


            TimeSpan df = e.AudioPosition;

            //double cPT = double.Parse(Math.Round(df.TotalSeconds).ToString() + "." +  Math.Round(df.TotalMilliseconds).ToString());
            double cPT = df.TotalSeconds;
            double xPT = (cPT * 30) / 100;
            xPT = xPT - (xPT * 13 / 136);
            cPT = (cPT - xPT);

            //string aTime = df.Seconds.ToString() + "." + df.Milliseconds.ToString() + "s";
            //string caTime = df.Seconds.ToString() + "." + (df.Milliseconds - 2).ToString() + "s";

            string aTime = String.Format("{0:0.########}", cPT) + "s";
            string caTime = String.Format("{0:0.########}", cPT) + "s";

            double lTime = cPT + xPT;

            last_time = String.Format("{0:0.########}", lTime) + "s";


            string spTxt = "";
            #region pro text
            int sPoint = e.CharacterPosition;

            if (curFullTxt.IndexOf(" ", sPoint) != -1)
            {
                int ePoint = curFullTxt.IndexOf(" ", sPoint);
                try
                {
                    if (curFullTxt.Substring(sPoint - 1, 1) == "\"")
                    {
                        sPoint = sPoint - 1;
                    }
                }
                catch { }
                spTxt = curFullTxt.Substring(sPoint, ePoint - sPoint);

            }
            else { spTxt = e.Text; }
            #endregion

            spTxt = spTxt.Replace("<", "&lt;");
            spTxt = spTxt.Replace(">", "&gt;");
            spTxt = spTxt.Replace("&", "&amp;");

            Regex wFnx = new Regex("\\s*(.+?)\\s", RegexOptions.IgnoreCase);
            string mtchTxt = "";
            MatchCollection mtchPoint = null;
            int mInt = 0;


            if (wFnx.Match(htLineTxt, spanPoint).Success)
            {
                mtchPoint = wFnx.Matches(htLineTxt, spanPoint);
                try
                {
                    if (mtchPoint[0].Groups[1].Value == spTxt)
                    {
                        mtchTxt = mtchPoint[0].Groups[1].Value;
                        mInt = 0;
                    }
                    else if (mtchPoint[1].Groups[1].Value == spTxt)
                    {
                        mtchTxt = mtchPoint[1].Groups[1].Value;
                        mInt = 1;
                    }
                    else if (mtchPoint[2].Groups[1].Value == spTxt)
                    {
                        mtchTxt = mtchPoint[2].Groups[1].Value;
                        mInt = 2;
                    }
                    else if (mtchPoint[3].Groups[1].Value == spTxt)
                    {
                        mtchTxt = mtchPoint[3].Groups[1].Value;
                        mInt = 3;
                    }
                    else if (mtchPoint[4].Groups[1].Value == spTxt)
                    {
                        mtchTxt = mtchPoint[4].Groups[1].Value;
                        mInt = 4;
                    }

                }
                catch { }

            }

            if (spTxt == mtchTxt)
            {
                int stPoint = mtchPoint[mInt].Groups[1].Index;
                int etPoint = stPoint + spTxt.Length;
                htLineTxt = htLineTxt.Insert(etPoint, "</span>");
                string spanidTxt = "<span id=\"w" + spanID.ToString() + "\">";
                htLineTxt = htLineTxt.Insert(stPoint, spanidTxt);


                spanPoint = etPoint + 7 + spanidTxt.Length;
            }

            smilTxt = smilTxt.Replace("[etime/]", caTime);

            string pTxt = "<par id=\"par" + spanID.ToString() + "\">\n";
            pTxt += "<text src=\"" + curPageHtml + "#w" + spanID.ToString() + "\" />\n";
            pTxt += "<audio src=\"../audio/" + curaudioFile + "\" clipBegin=\"" + aTime + "\" clipEnd=\"[etime/]\"/>\n";
            pTxt += "</par>\n";
            smilTxt += pTxt;

            spanID++;

            //probar update
            try
            {
                probar.Value += 1;
            }
            catch { }

        }


        public void convert_mp3(string wavPath)
        {

            string[] wavFiles = Directory.GetFiles(wavPath, "*.wav");
            foreach (string w in wavFiles)
            {
                string sPath = Path.GetDirectoryName(w);
                string fName = Path.GetFileNameWithoutExtension(w);
                #region write mp3
                WaveStream InStr = new WaveStream(w);
                try
                {
                    Mp3Writer writer = new Mp3Writer(new FileStream(sPath + "\\" + fName + ".mp3",
                                                        FileMode.Create), InStr.Format);
                    try
                    {
                        byte[] buff = new byte[writer.OptimalBufferSize];
                        int read = 0;
                        while ((read = InStr.Read(buff, 0, buff.Length)) > 0)
                        {
                            writer.Write(buff, 0, read);
                        }
                    }
                    finally
                    {
                        writer.Close();
                    }
                }
                finally
                {
                    InStr.Close();
                }

                #endregion
                File.Delete(w);
            }

        }
        public void SaveStreamToFile(string fileFullPath, Stream stream)
        {
            try
            {
                if (stream.Length == 0) return;
                using (FileStream fileStream = System.IO.File.Create(fileFullPath, (int)stream.Length))
                {
                    byte[] bytesInStream = new byte[stream.Length];
                    stream.Read(bytesInStream, 0, (int)bytesInStream.Length);
                    fileStream.Write(bytesInStream, 0, bytesInStream.Length);
                }
            }
            catch { }
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            try
            {
                Application.Exit();
            }
            catch (Exception erd) {
                gCls.show_error(erd.Message.ToString());
                return;
            }
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            browse_savemp3();
        }

        public void browse_savemp3() {
            try
            {
                SaveFileDialog sFld = new SaveFileDialog();
                sFld.Title = "Select file path for save MP3 File";
                sFld.Filter = "MP3 File|*.mp3";
                sFld.ShowDialog();
                if (sFld.FileName != "") {
                    txt_savefile.Text = sFld.FileName;
                }

            }
            catch(Exception erd) {
                gCls.show_error(erd.Message.ToString());
                return;
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            try
            {
                gCls.show_message(Application.ProductName + " " + Application.ProductVersion + "\nSend your Feedbacks to : vickypatel2020@gmail.com\n");
            }
            catch { }
        }

    }
}
