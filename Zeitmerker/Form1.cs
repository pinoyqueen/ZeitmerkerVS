using   CsvHelper;
using CsvHelper.Configuration;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

namespace Zeitmerker
{
    public partial class Form1 : Form
    {
        List<Prozesse> _prozesseList = new List<Prozesse>();
        List<Boolean> _clickedList = new List<Boolean>();
        List<Stopwatch> _stopwatchList = new List<Stopwatch>();
        List<Button> _buttonList = new List<Button>();
        List<Label> _labelList = new List<Label>();
        List<TimeSpan> _offsetList = new List<TimeSpan>();

        static System.Windows.Forms.Timer myTimer = new System.Windows.Forms.Timer();

        string XMLFilePath = Environment.CurrentDirectory + "\\prozesse.xml";
        string CSVFilePath = "prozesse.csv";

        public Form1()
        {
            InitializeComponent();

            myTimer.Tick += new EventHandler(TimerEventProcessor);
            myTimer.Interval = 1000;
            myTimer.Start();

            #region XML Datei erstellen
            ///// Projektnamen
            //for (int j = 0; j < 6; j++)
            //{
            //    _prozesseList.Add(new Prozesse() { Names = "Projekt" + j });
            //}

            ///// alle Prozesse in _prozesseList in einer XML-Datei erstellen 
            //try
            //{
            //    XmlSerializer x = new XmlSerializer(_prozesseList.GetType());
            //    using (FileStream fs = new FileStream(Environment.CurrentDirectory + "\\prozesse.xml", FileMode.Create, FileAccess.Write))
            //    {
            //        x.Serialize(fs, _prozesseList);
            //        MessageBox.Show("created");
            //    }
            //}
            //catch (Exception ex)
            //{
            //    throw new Exception("Fehler beim Speichern der XML Datei. " + ex.Message);

            //}
            #endregion
           
            #region XML Datei lesen
            /// XML - Datei lesen
            try
            {
                XmlSerializer x = new XmlSerializer(_prozesseList.GetType());
                using (FileStream fs = new FileStream(XMLFilePath, FileMode.Open, FileAccess.Read))
                {
                    _prozesseList = (List<Prozesse>)x.Deserialize(fs);
                }

                foreach (Prozesse pro in _prozesseList)
                {
                    _offsetList.Add(TimeSpan.Parse(pro.Zeit));
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Fehler beim Laden der XML Datei " + ex.Message);

            }
            #endregion

            #region dynamische Buttons
            /// dynamisch Buttons und Labels erstellen für jede Prozesse/Aufgabe
            int left = 10;
            int top = 20;

            for (int l = 0; l < _prozesseList.Count(); l++)
            {
                
                Button btn = new Button()
                {
                    Text = _prozesseList[l].Names.ToString(),
                    Width = 96,
                    Height = 57,
                    Left = left,
                    Top = top
                };

                Boolean clicked = false;
                Stopwatch stopwatch = new Stopwatch();
                Label label = new Label()
                {

                    Width = 96,
                    Height = 57,
                    Top = btn.Top + btn.Height,
                    Left = btn.Left,
                    AutoSize = true
                };

                _clickedList.Add(clicked);
                _stopwatchList.Add(stopwatch);
                _buttonList.Add(btn);
                _labelList.Add(label);

                btn.MouseUp += (s, e) => {
                    if (e.Button == MouseButtons.Right) {reset(_buttonList.IndexOf(btn)); }
                    else { timerStartStop(_buttonList.IndexOf(btn), btn); }
                };

                this.Controls.Add(btn);
                this.Controls.Add(label);

                if (btn.Left + btn.Width + left >= this.Width)
                {
                    top += btn.Height + 20;
                    left = 10;
                } else left += + 10 + btn.Width;  
            }

            #endregion 

        }

        /// <summary>
        /// Zeigt die aktuelle Datum und Uhrzeit
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TimerEventProcessor(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            label1.Text = now.ToString("f");
        }

        /// <summary>
        /// Wenn der Button zuerst gedrückt: startet den Timer
        /// und nochmal, um Timer anzuhalten und die Dauer anzuzeigen
        /// </summary>
        /// <param name="index">Index der Button in der Liste, um andere Controls zuzuordnen</param>
        /// <param name="button">welche Button gedrückt wird</param>
        public void timerStartStop(int index, Button button)
        {
            Boolean clicked = _clickedList[index];
            Stopwatch stopwatch = _stopwatchList[index];

            if (clicked == false)
            {
                stopwatch.Start();
                _clickedList[index] = true;
                button.BackColor = Color.Green;
                _labelList[index].Text = "";
            }
            else
            {
                stopwatch.Stop();
                _clickedList[index] = false;
                button.BackColor = SystemColors.ButtonFace;
                button.UseVisualStyleBackColor = true;

                TimeSpan offset = new TimeSpan();
                offset = _offsetList[index];

                TimeSpan ts = stopwatch.Elapsed + offset;
                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}:{3:00}", ts.Days, ts.Hours, ts.Minutes, ts.Seconds);

                _prozesseList[index].Aktualisiert = DateTime.Now.ToString("g");
                _prozesseList[index].Zeit = elapsedTime; 

                _labelList[index].Text = elapsedTime;
            }
        }

        /// <summary>
        /// Exportiert eine CSV Datei und aktualisert die XML Datei
        /// </summary>
        /// <param name="csvFile"></param>
        /// <param name="xmlFile"></param>
        /// <exception cref="Exception"></exception>
        public void Export(string csvFile, string xmlFile)
        {
            DialogResult dialogResult = MessageBox.Show("Wollen Sie die Datei exportieren?", "", MessageBoxButtons.OKCancel);
            if (dialogResult == DialogResult.OK)
            {
                foreach (Stopwatch stopwatch in _stopwatchList)
                {
                    stopwatch.Stop();
                }

                /// CSV Datei exportieren
                try
                {
                    using (StreamWriter sw = new StreamWriter(csvFile))
                    {
                        var config = new CsvConfiguration(CultureInfo.CurrentCulture) { Delimiter = ";", Encoding = Encoding.UTF8 };
                        using (var csv = new CsvWriter(sw, config))
                        {
                            csv.WriteRecords(_prozesseList);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Fehler beim Schreiben der Excel Datei " + ex.Message);
                }

                /// XML Datei aktualisieren
                try
                {
                    XmlSerializer x = new XmlSerializer(_prozesseList.GetType());
                    using (FileStream fs = new FileStream(xmlFile, FileMode.Create, FileAccess.Write))
                    {
                        x.Serialize(fs, _prozesseList);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Fehler beim Speichern der XML Datei. " + ex.Message);

                }
                this.Close();
            }

        }

        /// <summary>
        /// Setzt die Stopwatch zurück
        /// </summary>
        /// <param name="index"></param>
        public void reset(int index)
        {
            DialogResult dialogResult = MessageBox.Show("Wollen Sie die Zeit zurücksetzen?","", MessageBoxButtons.OKCancel);
            if (dialogResult == DialogResult.OK)
            {
                _stopwatchList[index].Reset();
                _prozesseList[index].Zeit = "00:00:00:00";
                _offsetList[index] = TimeSpan.Zero;
            }
        }

        /// <summary>
        /// Export button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {

            Export(CSVFilePath,XMLFilePath);
        }
    }

    public class Prozesse
    {
        // Name der Prozesse/Arbeitspaket
        public string Names { get; set; }
        // Die Dauer für jede Prozesse/Arbeitspaket
        public string Zeit { get; set; }
        // Wann der Button zuletzt aktiviert (aktuelle Datum und Uhrzeit)
        public string Aktualisiert { get; set; }
    }
}
