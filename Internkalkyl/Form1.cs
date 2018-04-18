using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;



namespace Internkalkyl
{
    public partial class Form1 : Form
    {
        private System.Drawing.Printing.PrintDocument printDocCaptureScreen = new System.Drawing.Printing.PrintDocument();
        private System.Drawing.Printing.PrintDocument printDocTextOnly = new System.Drawing.Printing.PrintDocument();

        public Form1()
        {
            InitializeComponent();

            Logg.toLog("Start");

            // Eventhandlers - designer only accepts one/type it seems or stuff fucks up
            // Milproduktionsbaserat
            myNumboxMilprodAntalMil.TextChanged += myNumboxKostnBransleOljaPerMil_TextChanged;
            myNumboxMilprodAntalMil.TextChanged += myNumboxDackskostnadMil_TextChanged;
            myNumboxMilprodAntalMil.TextChanged += myNumboxRantaRorelseKapitalKrMil_TextChanged;
            myNumboxMilprodAntalMil.TextChanged += myNumboxKostnadRepUnderhPerMil_TextChanged;
            myNumboxMilprodAntalMil.TextChanged += beraknaMilkostnad;

            // Beräkning av summa:
            myNumboxSummaPersKostn.TextChanged += beraknaSumma;
            myNumboxAvskrivningarSummaBelopp.TextChanged += beraknaSumma;
            myNumboxBerKalkylrantaAvskr.TextChanged += beraknaSumma;
            myNumboxBerKalkylrantaAvskr.TextChanged += beraknaSumma;
            myNumboxDrivmedelOlja.TextChanged += beraknaSumma;
            myNumboxDackskostnad.TextChanged += beraknaSumma;
            myNumboxOvrKostRantaRorligtKap.TextChanged += beraknaSumma;
            myNumboxReparationochUnderhall.TextChanged += beraknaSumma;
            myNumboxSummaOvrigaKostk.TextChanged += beraknaSumma;

            // när summan ändras
            myNumboxSumma.TextChanged += beraknaMilkostnad;


            // totalsumman
            myNumboxSumma.TextChanged += beraknaTotalt;
            myNumboxVinstRisk.TextChanged += beraknaTotalt;
            myNumboxAvgSidointkt.TextChanged += beraknaTotalt;

            // Beräkna vinst/risk
            myNumboxSumma.TextChanged += beraknaVinstRisk;
            myNumboxVinstRiskProc.TextChanged += beraknaVinstRisk;

            // Milkostn ink vinst/risk
            myNumboxSumma.TextChanged += beraknaMilkostnInkVinstRisk;
            myNumboxVinstRisk.TextChanged += beraknaMilkostnInkVinstRisk;
            myNumboxMilprodAntalMil.TextChanged += beraknaMilkostnInkVinstRisk;

            // Milkostnad efter sidointäkt
            myNumboxTotalt.TextChanged += beraknaMilkostnESidointakt;
            myNumboxMilprodAntalMil.TextChanged += beraknaMilkostnESidointakt;

            // Reducerad total pga samlastning
            myNumboxTotalt.TextChanged += myNumboxFyllnadsgradProc_TextChanged;

            // Print
            printDocCaptureScreen.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(printDocumentCaptureScreen_PrintPage);
            printDocTextOnly.PrintPage += new System.Drawing.Printing.PrintPageEventHandler(printTextOnly_PrintPage);

        }



        #region Avskrivningar

        /// <summary>
        /// Beräknar kostnad/avskrivningsår för lastbil
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void myNumboxLastbAnskVarde_Ar_TextChanged(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(myNumboxLastbAnskVarde.Text))
            {
                myNumboxLastbArligAvskrivning.Text = string.Empty;
                return;
            }

            if (String.IsNullOrWhiteSpace(myNumboxLastbAvskrtidAr.Text))
            {
                myNumboxLastbArligAvskrivning.Text = string.Empty;
                return;
            }

            double value = Convert.ToDouble(myNumboxLastbAnskVarde.Text);
            double years = Convert.ToDouble(myNumboxLastbAvskrtidAr.Text);
            double quotient = value / years;

            Logg.toLog("A change Value: " + value + " years " + years + " quotient: " + value / years);

            myNumboxLastbArligAvskrivning.Text = quotient.ToString("F2");

        }

        /// <summary>
        /// Beräknar kostnad/avskrivningsår för släp
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void myNumboxSlapAnskVarde_Ar_TextChanged(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(myNumboxSlapAnskVarde.Text))
            {
                myNumboxSlapArligAvskrivning.Text = string.Empty;
                return;
            }

            if (String.IsNullOrWhiteSpace(myNumboxSlapAvskrtidAr.Text))
            {
                myNumboxSlapArligAvskrivning.Text = string.Empty;
                return;
            }

            double value = Convert.ToDouble(myNumboxSlapAnskVarde.Text);
            double years = Convert.ToDouble(myNumboxSlapAvskrtidAr.Text);
            double quotient = value / years;

            Logg.toLog("Value: " + value + " years " + years + " quotient: " + value / years);

            myNumboxSlapArligAvskrivning.Text = quotient.ToString("F2");
        }

        /// <summary>
        /// Beräknar kostnad/avskrivningsår för kylaggregat
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void myNumboxKylaggAnskVarde_Ar_TextChanged(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(myNumboxKylaggAnskVarde.Text))
            {
                myNumboxKylaggArligAvskrivning.Text = string.Empty;
                return;
            }

            if (String.IsNullOrWhiteSpace(myNumboxKylaggAvskrtidAr.Text))
            {
                myNumboxKylaggArligAvskrivning.Text = string.Empty;
                return;
            }

            double value = Convert.ToDouble(myNumboxKylaggAnskVarde.Text);
            double years = Convert.ToDouble(myNumboxKylaggAvskrtidAr.Text);
            double quotient = value / years;

            Logg.toLog("Value: " + value + " years " + years + " quotient: " + value / years);

            myNumboxKylaggArligAvskrivning.Text = quotient.ToString("F2");
        }


        private void berSummaAvskrivningar_TextChanged(object sender, EventArgs e)
        {
            double inskvarde_lastbil;
            double inskvarde_slap;
            double inskvarde_kylagg;
            double avskr_ar_lastbil;
            double avskr_ar_slap;
            double avskr_ar_kylagg;


            try
            {
                inskvarde_lastbil = Convert.ToDouble(myNumboxLastbAnskVarde.Text);
            }
            catch
            {
                inskvarde_lastbil = 0;
            }
            try
            {
                inskvarde_slap = Convert.ToDouble(myNumboxSlapAnskVarde.Text);
            }
            catch
            {
                inskvarde_slap = 0;
            }
            try
            {
                inskvarde_kylagg = Convert.ToDouble(myNumboxKylaggAnskVarde.Text);
            }
            catch
            {
                inskvarde_kylagg = 0;
            }
            try
            {
                avskr_ar_lastbil = Convert.ToDouble(myNumboxLastbAvskrtidAr.Text);
            }
            catch
            {
                inskvarde_lastbil = 0;
                avskr_ar_lastbil = 1;
            }
            try
            {
                avskr_ar_slap = Convert.ToDouble(myNumboxSlapAvskrtidAr.Text);
            }
            catch
            {
                inskvarde_slap = 0;
                avskr_ar_slap = 1;
            }
            try
            {
                avskr_ar_kylagg = Convert.ToDouble(myNumboxKylaggAvskrtidAr.Text);
            }
            catch
            {
                inskvarde_kylagg = 0;
                avskr_ar_kylagg = 1;
            }

            double summa_avskr = (inskvarde_lastbil / avskr_ar_lastbil + inskvarde_kylagg / avskr_ar_kylagg + inskvarde_slap / avskr_ar_slap);

            myNumboxAvskrivningarSummaBelopp.Text = summa_avskr.ToString("F2");
        }



        #endregion // Avskrivningar




        #region Forare

        private void myNumboxAntForartimmar_TextChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(myNumboxAntForartimmar.Text))
            {
                myNumboxKvotFullTjanst.Text = string.Empty;
                return;
            }

            double antal_timmar = Convert.ToDouble(myNumboxAntForartimmar.Text);
            double kvotfulltjanst = antal_timmar / 1650;

            myNumboxKvotFullTjanst.Text = kvotfulltjanst.ToString("F2");
        }


        private void myNumboxForarkostnadPerTimme_TextChanged(object sender, EventArgs e)
        {
            // Lönekostnad förare
            if (string.IsNullOrWhiteSpace(myNumboxForarkostnadPerTimme.Text) || string.IsNullOrWhiteSpace(myNumboxAntForartimmar.Text))
            {
                myNumboxLoneKostnForare.Text = string.Empty;
                myNumboxSocKostnForare.Text = string.Empty;
                return;
            }

            double antal_timmar = Convert.ToDouble(myNumboxAntForartimmar.Text);
            double kostnad_timme = Convert.ToDouble(myNumboxForarkostnadPerTimme.Text);

            double lonekostn_forare = antal_timmar * kostnad_timme;

            myNumboxLoneKostnForare.Text = lonekostn_forare.ToString("F2");

            // Sociala kostnader förare
            if (string.IsNullOrWhiteSpace(myNumboxSocialaKostnaderProcent.Text))
            {
                myNumboxSocKostnForare.Text = string.Empty;
                return;
            }

            double soc_kostn_proc = Convert.ToDouble(myNumboxSocialaKostnaderProcent.Text);
            double soc_kostn_for = lonekostn_forare * soc_kostn_proc / 100;

            myNumboxSocKostnForare.Text = soc_kostn_for.ToString("F2");

        }

        private void myNumboxSumWorkerCosts_TextChanged(object sender, EventArgs e)
        {
            double admkostn;
            double lonekostn;
            double sockostn;
            double traktamente;
            double overnattn;
            try
            {
                admkostn = Convert.ToDouble(myNumboxAdmKostInkSoc.Text);
            }
            catch
            {
                admkostn = 0;
            }
            try
            {
                lonekostn = Convert.ToDouble(myNumboxLoneKostnForare.Text);
            }
            catch
            {
                lonekostn = 0;
            }
            try
            {
                sockostn = Convert.ToDouble(myNumboxSocKostnForare.Text);
            }
            catch
            {
                sockostn = 0;
            }
            try
            {
                traktamente = Convert.ToDouble(myNumboxTraktamenten.Text);
            }
            catch
            {
                traktamente = 0;
            }
            try
            {
                overnattn = Convert.ToDouble(myNumboxOvernattningskostnader.Text);
            }
            catch
            {
                overnattn = 0;
            }

            double arbkostn = admkostn + lonekostn + sockostn + traktamente + overnattn;

            myNumboxSummaPersKostn.Text = arbkostn.ToString("F2");

        }

        #endregion // Forare

        #region Övriga kostnader
        private void myNumboxSumOvrKostnKostnader_TextChanged(object sender, EventArgs e)
        {
            double lokalhyra;
            double fordonsbikostn;
            double uppvarmn_ext;
            double telefon;
            double storforforsakr;
            double fordonsskatt;
            double reservfordkostn;

            try
            {
                lokalhyra = Convert.ToDouble(myNumboxLokalhyrorKostnader.Text);
            }
            catch
            {
                lokalhyra = 0;
            }
            try
            {
                fordonsbikostn = Convert.ToDouble(myNumboxFordonsbikostnader.Text);
            }
            catch
            {
                fordonsbikostn = 0;
            }
            try
            {
                uppvarmn_ext = Convert.ToDouble(myNumboxUppvarmingExtern.Text);
            }
            catch
            {
                uppvarmn_ext = 0;
            }
            try
            {
                telefon = Convert.ToDouble(myNumboxTelefon.Text);
            }
            catch
            {
                telefon = 0;
            }
            try
            {
                storforforsakr = Convert.ToDouble(myNumboxStorforetagsforsakring.Text);
            }
            catch
            {
                storforforsakr = 0;
            }
            try
            {
                fordonsskatt = Convert.ToDouble(myNumboxFordonskatt.Text);
            }
            catch
            {
                fordonsskatt = 0;
            }
            try
            {
                reservfordkostn = Convert.ToDouble(myNumboxReservfordonskostnad.Text);
            }
            catch
            {
                reservfordkostn = 0;
            }

            double summaOvrigt = lokalhyra + fordonsbikostn + uppvarmn_ext + telefon + storforforsakr + fordonsskatt + reservfordkostn;

            myNumboxSummaOvrigaKostk.Text = summaOvrigt.ToString("F2");

        }


        #endregion // Övriga kostnader

        private void berKalkylrantaPaAvskrivningar_TextChanged(object sender, EventArgs e)
        {
            double kalkylranta_proc;
            double summa_avskr;

            try
            {
                kalkylranta_proc = Convert.ToDouble(myNumboxKalkylrantaProc.Text);
            }
            catch
            {
                kalkylranta_proc = 0;
            }
            try
            {
                summa_avskr = Convert.ToDouble(myNumboxAvskrivningarSummaBelopp.Text);
            }
            catch
            {
                summa_avskr = 0;
            }

            double kalkylranta_avskr = kalkylranta_proc / 100 * summa_avskr;

            myNumboxBerKalkylrantaAvskr.Text = kalkylranta_avskr.ToString("F2");

        }

        #region milbaserade kostnader
        private void myNumboxRantaRorelseKapitalKrMil_TextChanged(object sender, EventArgs e)
        {
            double milproduktion;
            double ranta_rorelsekap_mil;

            try
            {
                milproduktion = Convert.ToDouble(myNumboxMilprodAntalMil.Text);
            }
            catch
            {
                milproduktion = 0;
            }
            try
            {
                ranta_rorelsekap_mil = Convert.ToDouble(myNumboxRantaRorelseKapitalKrMil.Text);
            }
            catch
            {
                ranta_rorelsekap_mil = 0;
            }

            double ranta_rorelsekap = ranta_rorelsekap_mil * milproduktion;

            myNumboxOvrKostRantaRorligtKap.Text = ranta_rorelsekap.ToString("F2");

        }

        private void myNumboxKostnBransleOljaPerMil_TextChanged(object sender, EventArgs e)
        {
            double milproduktion;
            double kostn_bransle_olja_per_mil;

            try
            {
                milproduktion = Convert.ToDouble(myNumboxMilprodAntalMil.Text);
            }
            catch
            {
                milproduktion = 0;
            }
            try
            {
                kostn_bransle_olja_per_mil = Convert.ToDouble(myNumboxKostnBransleOljaPerMil.Text);
            }
            catch
            {
                kostn_bransle_olja_per_mil = 0;
            }

            double bransle_olja = kostn_bransle_olja_per_mil * milproduktion;
            myNumboxDrivmedelOlja.Text = bransle_olja.ToString("F2");

        }

        private void myNumboxDackskostnadMil_TextChanged(object sender, EventArgs e)
        {
            double milproduktion;
            double dackskostn_mil;

            try
            {
                milproduktion = Convert.ToDouble(myNumboxMilprodAntalMil.Text);
            }
            catch
            {
                milproduktion = 0;
            }
            try
            {
                dackskostn_mil = Convert.ToDouble(myNumboxDackskostnadMil.Text);
            }
            catch
            {
                dackskostn_mil = 0;
            }

            double dackskostn = dackskostn_mil * milproduktion;
            myNumboxDackskostnad.Text = dackskostn.ToString("F2");
        }

        private void myNumboxKostnadRepUnderhPerMil_TextChanged(object sender, EventArgs e)
        {
            double milproduktion;
            double repunderhall_mil;

            try
            {
                milproduktion = Convert.ToDouble(myNumboxMilprodAntalMil.Text);
            }
            catch
            {
                milproduktion = 0;
            }
            try
            {
                repunderhall_mil = Convert.ToDouble(myNumboxKostnadRepUnderhPerMil.Text);
            }
            catch
            {
                repunderhall_mil = 0;
            }

            double repunderhall = repunderhall_mil * milproduktion;
            myNumboxReparationochUnderhall.Text = repunderhall.ToString("F2");
        }


        #endregion // milbaserade kostnader


        #region Summering
        private void beraknaMilkostnESidointakt(object sender, EventArgs e)
        {
            double milproduktion;
            double totalt;

            try
            {
                milproduktion = Convert.ToDouble(myNumboxMilprodAntalMil.Text);
            }
            catch
            {
                milproduktion = 0;
            }
            try
            {
                totalt = Convert.ToDouble(myNumboxTotalt.Text);
            }
            catch
            {
                totalt = 0;
            }

            double milkostnad_efter_sidointakt;

            if (milproduktion == 0)
                milkostnad_efter_sidointakt = 0;
            else
                milkostnad_efter_sidointakt = totalt / milproduktion;

            myNumboxSumMilkostEftSidoint.Text = milkostnad_efter_sidointakt.ToString("F2");
        }

        private void beraknaMilkostnad(object sender, EventArgs e)
        {
            double milproduktion;
            double summa;

            try
            {
                milproduktion = Convert.ToDouble(myNumboxMilprodAntalMil.Text);
            }
            catch
            {
                milproduktion = 0;
            }
            try
            {
                summa = Convert.ToDouble(myNumboxSumma.Text);
            }
            catch
            {
                summa = 0;
            }
            double milkostnad;

            if (milproduktion == 0)
                milkostnad = 0;
            else
                milkostnad = summa / milproduktion;

            myNumboxSumMilkostn.Text = milkostnad.ToString("F2");

        }

        private void beraknaSumma(object sender, EventArgs e)
        {
            double personalkostnader;
            double arliga_avskrivningar;
            double kalkylranta_avskrivningar;
            double drivmedel_o_olja;
            double dackskostnader;
            double ranta_rorligt_kapital;
            double reparation_o_underhall;
            double ovriga_kostnader;

            try
            {
                personalkostnader = Convert.ToDouble(myNumboxSummaPersKostn.Text);
            }
            catch
            {
                personalkostnader = 0;
            }
            try
            {
                arliga_avskrivningar = Convert.ToDouble(myNumboxAvskrivningarSummaBelopp.Text);
            }
            catch
            {
                arliga_avskrivningar = 0;
            }
            try
            {
                kalkylranta_avskrivningar = Convert.ToDouble(myNumboxBerKalkylrantaAvskr.Text);
            }
            catch
            {
                kalkylranta_avskrivningar = 0;
            }
            try
            {
                drivmedel_o_olja = Convert.ToDouble(myNumboxDrivmedelOlja.Text);
            }
            catch
            {
                drivmedel_o_olja = 0;
            }
            try
            {
                dackskostnader = Convert.ToDouble(myNumboxDackskostnad.Text);
            }
            catch
            {
                dackskostnader = 0;
            }
            try
            {
                ranta_rorligt_kapital = Convert.ToDouble(myNumboxOvrKostRantaRorligtKap.Text);
            }
            catch
            {
                ranta_rorligt_kapital = 0;
            }
            try
            {
                reparation_o_underhall = Convert.ToDouble(myNumboxReparationochUnderhall.Text);
            }
            catch
            {
                reparation_o_underhall = 0;
            }
            try
            {
                ovriga_kostnader = Convert.ToDouble(myNumboxSummaOvrigaKostk.Text);
            }
            catch
            {
                ovriga_kostnader = 0;
            }

            double summa = personalkostnader + arliga_avskrivningar + kalkylranta_avskrivningar + drivmedel_o_olja + dackskostnader + ranta_rorligt_kapital + reparation_o_underhall + ovriga_kostnader;

            myNumboxSumma.Text = summa.ToString("F2");

        }

        private void beraknaTotalt(object sender, EventArgs e)
        {
            double summa;
            double vinst_risk;
            double avg_sidointakt;

            try
            {
                summa = Convert.ToDouble(myNumboxSumma.Text);
            }
            catch
            {
                summa = 0;
            }
            try
            {
                vinst_risk = Convert.ToDouble(myNumboxVinstRisk.Text);
            }
            catch
            {
                vinst_risk = 0;
            }
            try
            {
                avg_sidointakt = Convert.ToDouble(myNumboxAvgSidointkt.Text);
            }
            catch
            {
                avg_sidointakt = 0;
            }

            double totalt = summa + vinst_risk - avg_sidointakt;
            myNumboxTotalt.Text = totalt.ToString("F2");
        }

        private void beraknaVinstRisk(object sender, EventArgs e)
        {
            double summa;
            double onsk_vinst_risk_proc;


            try
            {
                summa = Convert.ToDouble(myNumboxSumma.Text);
            }
            catch
            {
                summa = 0;
            }
            try
            {
                onsk_vinst_risk_proc = Convert.ToDouble(myNumboxVinstRiskProc.Text);
            }
            catch
            {
                onsk_vinst_risk_proc = 0;
            }

            double vinst_risk = onsk_vinst_risk_proc / 100 * summa;

            myNumboxVinstRisk.Text = vinst_risk.ToString("F2");
        }


        private void beraknaMilkostnInkVinstRisk(object sender, EventArgs e)
        {
            double summa;
            double vinst_risk;
            double milproduktion;

            try
            {
                summa = Convert.ToDouble(myNumboxSumma.Text);
            }
            catch
            {
                summa = 0;
            }
            try
            {
                vinst_risk = Convert.ToDouble(myNumboxVinstRisk.Text);
            }
            catch
            {
                vinst_risk = 0;
            }
            try
            {
                milproduktion = Convert.ToDouble(myNumboxMilprodAntalMil.Text);
            }
            catch
            {
                milproduktion = 0;
            }
            double milkostnad_ink_vinst_risk;

            if (milproduktion == 0)
                milkostnad_ink_vinst_risk = 0;
            else
                milkostnad_ink_vinst_risk = (summa + vinst_risk) / milproduktion;

            myNumboxSumMilkostInkVinstRisk.Text = milkostnad_ink_vinst_risk.ToString("F2");
        }

        #endregion // Summering

        #region Samlasting

        private void myNumboxFyllnadsgradProc_TextChanged(object sender, EventArgs e)
        {
            double total;
            double samlastning_proc;

            try
            {
                total = Convert.ToDouble(myNumboxTotalt.Text);
            }
            catch
            {
                total = 0;
            }
            try
            {
                samlastning_proc = Convert.ToDouble(myNumboxFyllnadsgradProc.Text);
            }
            catch
            {
                samlastning_proc = 0;
            }

            double samlastning_reducerad_total = total * samlastning_proc / 100;

            myNumboxReduceradTotalSamlastning.Text = samlastning_reducerad_total.ToString("F2");


        }


        #endregion // Samlastning

        #region Sparaknapp

        private StringBuilder recursiveTextStringFinder(System.Windows.Forms.Control.ControlCollection cc, StringBuilder text)
        {
            GroupBox groupbox;
            TextBox textbox;

            foreach (Control C in cc)
            {
                try
                {
                    groupbox = (GroupBox)C;

                    text = (recursiveTextStringFinder(groupbox.Controls, text));
                }
                catch { } // do nothing

                try
                {
                    textbox = (TextBox)C;

                    text.AppendLine(textbox.Name + " " + textbox.Text);
                }
                catch { } // do nothing
            }

            return text;
        }

        private void buttonSpara_Click(object sender, EventArgs e)
        {
            SaveFileDialog savefiledialog = new SaveFileDialog();
            
            savefiledialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            savefiledialog.RestoreDirectory = true;
            savefiledialog.CheckFileExists = false;
            savefiledialog.CheckPathExists = true;

            if (string.IsNullOrEmpty(textBoxMainFritext.Text))
                savefiledialog.FileName = "";
            else
                savefiledialog.FileName = textBoxMainFritext.Text;

            if(savefiledialog.ShowDialog() == DialogResult.OK)
            {
                System.IO.StreamWriter stream = new System.IO.StreamWriter(savefiledialog.FileName);

                StringBuilder text = new StringBuilder();

                stream.WriteLine("// This file was created using the program Internkalkyl, copyright Northern Communications 2017");

                stream.WriteLine(recursiveTextStringFinder(this.Controls, text));

                stream.Close();
                                
            }
        }
        #endregion //Sparaknapp

        #region Hämtaknapp

        private void recursiveControlPopulator(System.Windows.Forms.Control.ControlCollection cc, String[] settings)
        {
            GroupBox groupbox;
            TextBox textbox;

            foreach (Control C in cc)
            {
                try
                {
                    groupbox = (GroupBox)C;

                    recursiveControlPopulator(groupbox.Controls, settings);
                }
                catch { } // do nothing

                try
                {
                    textbox = (TextBox)C;

                    string current = Array.Find(settings, s => s.Contains(textbox.Name));
                    
                    if (!String.IsNullOrEmpty(current))
                        textbox.Text = current.Substring(textbox.Name.Length + 1, current.Length - textbox.Name.Length-1);
                   
                }
                catch { } // do nothing
            }
        }

        private void buttonHamta_Click(object sender, EventArgs e)
        {
            OpenFileDialog opf = new OpenFileDialog();

            opf.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
            opf.RestoreDirectory = true;

            if (opf.ShowDialog() == DialogResult.OK)
            {
                string[] settings = File.ReadAllLines(opf.FileName);

                recursiveControlPopulator(this.Controls, settings);
            }

            myNumboxForarkostnadPerTimme_TextChanged(this, new EventArgs());
        }

        #endregion // Hämtaknapp

        #region Printknapp
        private System.Text.RegularExpressions.Regex reasonableFormat = new System.Text.RegularExpressions.Regex(@"^(\d+)([xyz]*)([biulcrtg]*)$");
        // xyz - column 1,2 or 3
        // biulcrt - bold, italics, underline, left, center, right adjust, title (larger size text + extra space before)
        
        private int recursiveControlListMaker(Control.ControlCollection CC)
        {
            GroupBox groupbox;
            TextBox textbox;
            Label label;
            string formatting;
            int column, index;
            int numRowsTotal = 0;
            int numRowsGroupBox = 0;
            System.Text.RegularExpressions.Match formattingMatch;

            foreach (Control C in CC)
            {
                try
                {
                    if (reasonableFormat.IsMatch((string)C.Tag))
                    {
                        formattingMatch = reasonableFormat.Match((string)C.Tag);

                        index = Convert.ToInt32(formattingMatch.Groups[1].Value);

                        try
                        {
                            column = Convert.ToInt32((formattingMatch.Groups[2].Value.ToCharArray())[0] - 'w');
                        }
                        catch
                        {
                            column = 0;
                        }
                        try
                        {
                            formatting = formattingMatch.Groups[3].ToString();
                        }
                        catch
                        {
                            formatting = "";
                        }
                    }
                    else
                    {
                        index = -1;
                        formatting = "";
                        column = 0;
                    }
                }
                catch
                {
                    index = -1;
                    formatting = "";
                    column = 0;
                }
                    
                try
                {
                    groupbox = (GroupBox)C;

                    try
                    {
                        numRowsGroupBox = recursiveControlListMaker(groupbox.Controls);
                        numRowsTotal += numRowsGroupBox;
                        numRowsTotal++;

                        try
                        {
                            globallistOfLabelandValue[index, 0] = groupbox.Text;
                            globallistOfLabelandValue[index, 4] = formatting + 't';
                            globallistOfLabelandValue[index, 5] = Convert.ToString(numRowsGroupBox);
                        }
                        catch { }
                    }
                    catch (Exception ex)
                    {
                        Logg.toLog("Exception " + ex.Message);
                    }
                    continue;

                }
                catch { } // do nothing

                try
                {
                    textbox = (TextBox)C;

                    if (column == 0)
                    {
                        globallistOfLabelandValue[index, 1] = textbox.Text;
                    }
                    else
                        globallistOfLabelandValue[index, column] = textbox.Text;

                    globallistOfLabelandValue[index, 4] = formatting;

                    if (column == 0 || column == 1)
                        numRowsTotal++;

                    continue;
                }
                catch { } // do nothing

                try
                {
                    label = (Label)C;

                    globallistOfLabelandValue[index, column] = label.Text;
                    globallistOfLabelandValue[index, 4] = formatting;

                    continue;
                }
                catch (Exception ex)
                {
                    Logg.toLog("Exception: " + ex.Message);

                } // do nothing
                
            }
            return numRowsTotal;
        }

        private System.Text.RegularExpressions.Regex onlyDigits = new System.Text.RegularExpressions.Regex(@"^\d+$");
        private void printTextOnly_PrintPage(System.Object sender, System.Drawing.Printing.PrintPageEventArgs ev)
        {
            
            float leftMargin = ev.MarginBounds.Left;
            float topMargin = ev.MarginBounds.Top;
            float tab = (ev.MarginBounds.Right - ev.MarginBounds.Left) / 5;
            float tab1 = tab * 3;
            float tab2 = tab1 + tab;
            float tab3 = tab2 + tab;

            Font printFont = new Font("Arial", 10);
            Font printFontBold = new Font("Arial", 10, FontStyle.Bold);
            Font printFontTitle = new Font("Arial", 12);
            Font currentFont = printFont;

            offsetOtherPages = topMargin * numberOfPages + ev.MarginBounds.Bottom * (numberOfPages - 1);
            pagePosForTextPrint = topMargin;

            while (pagePosForTextPrint < ev.MarginBounds.Bottom && countForTextPrint < globallistOfLabelandValue.GetLength(0))
            {
                if (!string.IsNullOrEmpty(globallistOfLabelandValue[countForTextPrint, 4]))
                {
                    if (globallistOfLabelandValue[countForTextPrint, 4].Contains('b'))
                    {
                        currentFont = printFontBold;
                    }
                    else if (globallistOfLabelandValue[countForTextPrint, 4].Contains('t'))
                    {
                        try
                        {
                            float heightOfParagraph = (Convert.ToInt32(globallistOfLabelandValue[countForTextPrint, 5]) + 2) * printFontBold.GetHeight(ev.Graphics) + printFontTitle.GetHeight();

                            if (pagePosForTextPrint + heightOfParagraph > ev.MarginBounds.Bottom)
                            {
                                ev.HasMorePages = true;
                                return;
                            }

                        }
                        catch { }

                        pagePosForTextPrint += printFont.GetHeight(ev.Graphics);
                        currentFont = printFontTitle;
                    }
                }
                else
                {
                    currentFont = printFont;
                }


                ev.Graphics.DrawString(globallistOfLabelandValue[countForTextPrint, 0], currentFont, Brushes.Black, leftMargin, pagePosForTextPrint, new StringFormat());
 
                if (globallistOfLabelandValue[countForTextPrint, 1] != null)
                {
                    if (string.IsNullOrEmpty(globallistOfLabelandValue[countForTextPrint, 0]) && (globallistOfLabelandValue[countForTextPrint, 4].Contains('t') || globallistOfLabelandValue[countForTextPrint, 4].Contains('l')))
                    {
                        ev.Graphics.DrawString(globallistOfLabelandValue[countForTextPrint, 1], currentFont, Brushes.Black, new RectangleF(leftMargin, pagePosForTextPrint, ev.MarginBounds.Right - leftMargin, ev.MarginBounds.Bottom - pagePosForTextPrint), new StringFormat());

                        SizeF size = ev.Graphics.MeasureString(globallistOfLabelandValue[countForTextPrint, 1], currentFont);

                        pagePosForTextPrint += size.Height + printFont.Height;
                    }
                    else
                    {
                        ev.Graphics.DrawString(globallistOfLabelandValue[countForTextPrint, 1], currentFont, Brushes.Black, tab1, pagePosForTextPrint, new StringFormat());
                    }
                }

                if (globallistOfLabelandValue[countForTextPrint, 2] != null)
                {
                    ev.Graphics.DrawString(globallistOfLabelandValue[countForTextPrint, 2], currentFont, Brushes.Black, tab2, pagePosForTextPrint, new StringFormat());
                }

                if (globallistOfLabelandValue[countForTextPrint, 3] != null)
                {
                    ev.Graphics.DrawString(globallistOfLabelandValue[countForTextPrint, 3], currentFont, Brushes.Black, tab3, pagePosForTextPrint, new StringFormat());
                }

                countForTextPrint++;
                pagePosForTextPrint += currentFont.GetHeight(ev.Graphics);
            }

            if (countForTextPrint < globallistOfLabelandValue.GetLength(0))
            {
                ev.HasMorePages = true;
                numberOfPages++;              
            }
            else
                ev.HasMorePages = false;

        }

        private string[,] globallistOfLabelandValue;
        private float pagePosForTextPrint;
        private float offsetOtherPages;
        private int countForTextPrint;
        private int numberOfPages;

        private void buttonSkrivUtEndText_Click(object sender, EventArgs e)
        {
            globallistOfLabelandValue = new string[55, 6];

            int numRows = recursiveControlListMaker(Controls);

            countForTextPrint = 0;
            offsetOtherPages = 0;
            numberOfPages = 1;

            PrintDialog pdi = new PrintDialog();
            pdi.Document = printDocTextOnly;

            if (pdi.ShowDialog() == DialogResult.OK)
            {
                printDocTextOnly.Print();
            }
        }

        private Bitmap memoryImage;

        private void captureScreen()
        {
            Graphics myGraphics = this.CreateGraphics();
            Size s = this.Size;
            memoryImage = new Bitmap(s.Width, s.Height, myGraphics);

            Graphics memoryGraphics = Graphics.FromImage(memoryImage);
            memoryGraphics.CopyFromScreen(this.Location.X, this.Location.Y, 0, 0, s);
        }


        private void buttonSkrivUt_Click(object sender, EventArgs e)
        {
            captureScreen();

            PrintDialog pdi = new PrintDialog();
            pdi.Document = printDocCaptureScreen;

            if (pdi.ShowDialog() == DialogResult.OK)
            {
                printDocCaptureScreen.Print();
            }
        }

        //public static Bitmap ResizeImage(Image image, int width, int height)
        //{
        //    var destRect = new Rectangle(0, 0, width, height);
        //    var destImage = new Bitmap(width, height);

        //    destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

        //    using (var graphics = Graphics.FromImage(destImage))
        //    {
        //        graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
        //        graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
        //        graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
        //        graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
        //        graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

        //        using (var wrapMode = new System.Drawing.Imaging.ImageAttributes())
        //        {
        //            wrapMode.SetWrapMode(System.Drawing.Drawing2D.WrapMode.TileFlipXY);
        //            graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
        //        }
        //    }

        //    return destImage;
        //}

        private void printDocumentCaptureScreen_PrintPage(System.Object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

            float ratio_height_width_img = (float)memoryImage.Height / (float)memoryImage.Width;
            float new_width = e.MarginBounds.Right - e.MarginBounds.Left;

            e.Graphics.DrawImage(memoryImage, e.MarginBounds.Left, e.MarginBounds.Top, new_width, new_width * ratio_height_width_img);
        }

       

        #endregion // Printknapp


    }
}
