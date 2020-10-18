using System;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;

namespace CheckInGUI
{
    
    partial class NewTemplate
    {
        public static string fileName;

        // Moves focus to the next control on enter keypress.
        public void Control_KeyUp(object sender, KeyEventArgs e)
        {
            if ((e.KeyCode == Keys.Enter) || (e.KeyCode == Keys.Return || e.KeyCode == Keys.Tab))
            {
                this.SelectNextControl((System.Windows.Forms.Control)sender, true, true, true, true);
            }
        }

        //Clears text boxes and returns focus to the first box.
        private void Control_ReturnFocus()
        {
            if (persistPartNumbers == true)
            {
                NewModelNum.Clear();
                NewSize.Clear();
            }
            else
            {
                NewModelNum.Clear();
                NewVS.Clear();
                NewOR1.Clear();
                NewOR2.Clear();
                NewSize.Clear();
                NewChemical.Clear();
                NewExtraParts.Clear();
                NewExtraLabor.Clear();
                NewHTYear.Clear();
                New6YRYear.Clear();
            }
                NewModelNum.Focus();
        }

        // Designer.
        private System.ComponentModel.IContainer components = null;

        // Dispose of resources.
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        // Serializes Template objects to a new XML file with the model number as the file name.
        public void SerializeObject<T>(T serializableObject, string fileName)
        {
            if (serializableObject == null) { return; }
            else
            {
                XmlDocument xmlDocument = new XmlDocument();
                XmlSerializer serializer = new XmlSerializer(serializableObject.GetType());
                using (MemoryStream stream = new MemoryStream())
                {
                    serializer.Serialize(stream, serializableObject);
                    stream.Position = 0;
                    xmlDocument.Load(stream);
                    xmlDocument.Save(fileName);
                }
            }
        }

        // Grabs Template objects from their respective template files.
        public void TemplateBuilder()
        {
            TemplateBuilder template = new TemplateBuilder();
            
            template.templateModel = entryModel.ToUpper();
            template.templateVS = entryVS;
            template.templateOR1 = entryOR1;
            template.templateOR2 = entryOR2;
            template.templateSize = entrySize;
            template.templateChem = entryChemical;
            template.templateExParts = entryExtraParts;
            template.templateExLabor = entryExtraLabor;
            template.templateHTYear = entryHTYearInt;
            template.template6YRYear = entry6YearInt;
            template.templateCollar = entryCollar;
            template.templateCT = entryCT;

            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string dirPath = path + "/Check-In Templates/";
            DirectoryInfo di = Directory.CreateDirectory(dirPath);
            string fileName = dirPath + template.templateModel + " template.xml";

            SerializeObject<TemplateBuilder>(template, fileName);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.multiSizeBox = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.NewLabel = new System.Windows.Forms.Label();
            this.NewCTYesB = new System.Windows.Forms.RadioButton();
            this.NewCTNoB = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.NewCollar = new System.Windows.Forms.Label();
            this.NewCollarYesB = new System.Windows.Forms.RadioButton();
            this.NewCollarNoB = new System.Windows.Forms.RadioButton();
            this.SubmissionB = new System.Windows.Forms.Button();
            this.Title = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.New6YRYear = new System.Windows.Forms.TextBox();
            this.NewHTYear = new System.Windows.Forms.TextBox();
            this.NewExtraLabor = new System.Windows.Forms.TextBox();
            this.NewExtraParts = new System.Windows.Forms.TextBox();
            this.NewChemical = new System.Windows.Forms.TextBox();
            this.NewSize = new System.Windows.Forms.TextBox();
            this.NewOR2 = new System.Windows.Forms.TextBox();
            this.NewOR1 = new System.Windows.Forms.TextBox();
            this.NewVS = new System.Windows.Forms.TextBox();
            this.NewModelNum = new System.Windows.Forms.TextBox();
            this.panel1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.panel1.Controls.Add(this.multiSizeBox);
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Controls.Add(this.SubmissionB);
            this.panel1.Controls.Add(this.Title);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.New6YRYear);
            this.panel1.Controls.Add(this.NewHTYear);
            this.panel1.Controls.Add(this.NewExtraLabor);
            this.panel1.Controls.Add(this.NewExtraParts);
            this.panel1.Controls.Add(this.NewChemical);
            this.panel1.Controls.Add(this.NewSize);
            this.panel1.Controls.Add(this.NewOR2);
            this.panel1.Controls.Add(this.NewOR1);
            this.panel1.Controls.Add(this.NewVS);
            this.panel1.Controls.Add(this.NewModelNum);
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(744, 450);
            this.panel1.TabIndex = 0;
            // 
            // multiSizeBox
            // 
            this.multiSizeBox.AutoSize = true;
            this.multiSizeBox.Location = new System.Drawing.Point(315, 237);
            this.multiSizeBox.Name = "multiSizeBox";
            this.multiSizeBox.Size = new System.Drawing.Size(130, 17);
            this.multiSizeBox.TabIndex = 32;
            this.multiSizeBox.Text = "Persist Part Numbers?";
            this.multiSizeBox.UseVisualStyleBackColor = true;
            this.multiSizeBox.CheckedChanged += new System.EventHandler(this.multiSizeBox_CheckedChanged);
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.groupBox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.groupBox2.Controls.Add(this.NewLabel);
            this.groupBox2.Controls.Add(this.NewCTYesB);
            this.groupBox2.Controls.Add(this.NewCTNoB);
            this.groupBox2.Location = new System.Drawing.Point(406, 260);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(200, 100);
            this.groupBox2.TabIndex = 31;
            this.groupBox2.TabStop = false;
            // 
            // NewLabel
            // 
            this.NewLabel.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.NewLabel.AutoSize = true;
            this.NewLabel.Location = new System.Drawing.Point(63, 29);
            this.NewLabel.Name = "NewLabel";
            this.NewLabel.Size = new System.Drawing.Size(56, 13);
            this.NewLabel.TabIndex = 23;
            this.NewLabel.Text = "CT Label?";
            // 
            // NewCTYesB
            // 
            this.NewCTYesB.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.NewCTYesB.AutoSize = true;
            this.NewCTYesB.Location = new System.Drawing.Point(66, 45);
            this.NewCTYesB.Name = "NewCTYesB";
            this.NewCTYesB.Size = new System.Drawing.Size(43, 17);
            this.NewCTYesB.TabIndex = 26;
            this.NewCTYesB.TabStop = true;
            this.NewCTYesB.Text = "Yes";
            this.NewCTYesB.UseVisualStyleBackColor = true;
            this.NewCTYesB.CheckedChanged += new System.EventHandler(this.NewCTYesB_CheckedChanged);
            // 
            // NewCTNoB
            // 
            this.NewCTNoB.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.NewCTNoB.AutoSize = true;
            this.NewCTNoB.Location = new System.Drawing.Point(66, 68);
            this.NewCTNoB.Name = "NewCTNoB";
            this.NewCTNoB.Size = new System.Drawing.Size(39, 17);
            this.NewCTNoB.TabIndex = 27;
            this.NewCTNoB.TabStop = true;
            this.NewCTNoB.Text = "No";
            this.NewCTNoB.UseVisualStyleBackColor = true;
            this.NewCTNoB.CheckedChanged += new System.EventHandler(this.NewCTNoB_CheckedChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.groupBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.groupBox1.Controls.Add(this.NewCollar);
            this.groupBox1.Controls.Add(this.NewCollarYesB);
            this.groupBox1.Controls.Add(this.NewCollarNoB);
            this.groupBox1.Location = new System.Drawing.Point(149, 260);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(200, 100);
            this.groupBox1.TabIndex = 30;
            this.groupBox1.TabStop = false;
            // 
            // NewCollar
            // 
            this.NewCollar.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.NewCollar.AutoSize = true;
            this.NewCollar.Location = new System.Drawing.Point(65, 29);
            this.NewCollar.Name = "NewCollar";
            this.NewCollar.Size = new System.Drawing.Size(39, 13);
            this.NewCollar.TabIndex = 22;
            this.NewCollar.Text = "Collar?";
            // 
            // NewCollarYesB
            // 
            this.NewCollarYesB.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.NewCollarYesB.AutoSize = true;
            this.NewCollarYesB.Location = new System.Drawing.Point(61, 45);
            this.NewCollarYesB.Name = "NewCollarYesB";
            this.NewCollarYesB.Size = new System.Drawing.Size(43, 17);
            this.NewCollarYesB.TabIndex = 24;
            this.NewCollarYesB.TabStop = true;
            this.NewCollarYesB.Text = "Yes";
            this.NewCollarYesB.UseVisualStyleBackColor = true;
            this.NewCollarYesB.CheckedChanged += new System.EventHandler(this.NewCollarYesB_CheckedChanged);
            // 
            // NewCollarNoB
            // 
            this.NewCollarNoB.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.NewCollarNoB.AutoSize = true;
            this.NewCollarNoB.Location = new System.Drawing.Point(61, 68);
            this.NewCollarNoB.Name = "NewCollarNoB";
            this.NewCollarNoB.Size = new System.Drawing.Size(39, 17);
            this.NewCollarNoB.TabIndex = 25;
            this.NewCollarNoB.TabStop = true;
            this.NewCollarNoB.Text = "No";
            this.NewCollarNoB.UseVisualStyleBackColor = true;
            this.NewCollarNoB.CheckedChanged += new System.EventHandler(this.NewCollarNoB_CheckedChanged);
            // 
            // SubmissionB
            // 
            this.SubmissionB.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.SubmissionB.AutoSize = true;
            this.SubmissionB.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SubmissionB.Location = new System.Drawing.Point(337, 378);
            this.SubmissionB.Name = "SubmissionB";
            this.SubmissionB.Size = new System.Drawing.Size(78, 34);
            this.SubmissionB.TabIndex = 29;
            this.SubmissionB.Text = "&Submit";
            this.SubmissionB.UseVisualStyleBackColor = true;
            this.SubmissionB.Click += new System.EventHandler(this.SubmissionB_Click);
            // 
            // Title
            // 
            this.Title.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.Title.AutoSize = true;
            this.Title.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Title.Location = new System.Drawing.Point(42, 49);
            this.Title.Name = "Title";
            this.Title.Size = new System.Drawing.Size(662, 31);
            this.Title.TabIndex = 28;
            this.Title.Text = "Please enter the information for the new template.";
            // 
            // label10
            // 
            this.label10.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(522, 180);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(78, 13);
            this.label10.TabIndex = 21;
            this.label10.Text = "New 6YR Year";
            // 
            // label9
            // 
            this.label9.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(403, 180);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(72, 13);
            this.label9.TabIndex = 20;
            this.label9.Text = "New HT Year";
            // 
            // label8
            // 
            this.label8.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(281, 180);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(86, 13);
            this.label8.TabIndex = 19;
            this.label8.Text = "New Extra Labor";
            // 
            // label7
            // 
            this.label7.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(151, 180);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(83, 13);
            this.label7.TabIndex = 18;
            this.label7.Text = "New Extra Parts";
            // 
            // label6
            // 
            this.label6.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(641, 120);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(75, 13);
            this.label6.TabIndex = 17;
            this.label6.Text = "New Chemical";
            // 
            // label5
            // 
            this.label5.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(507, 117);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(106, 13);
            this.label5.TabIndex = 16;
            this.label5.Text = "New Size (#, L, or G)";
            // 
            // label4
            // 
            this.label4.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(403, 120);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(74, 13);
            this.label4.TabIndex = 15;
            this.label4.Text = "New O-Ring 2";
            // 
            // label3
            // 
            this.label3.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(281, 120);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(74, 13);
            this.label3.TabIndex = 14;
            this.label3.Text = "New O-Ring 1";
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(140, 120);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(86, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "New Valve Stem";
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 121);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 13);
            this.label1.TabIndex = 12;
            this.label1.Text = "New Model #";
            // 
            // New6YRYear
            // 
            this.New6YRYear.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.New6YRYear.Location = new System.Drawing.Point(510, 196);
            this.New6YRYear.Name = "New6YRYear";
            this.New6YRYear.Size = new System.Drawing.Size(100, 20);
            this.New6YRYear.TabIndex = 9;
            this.New6YRYear.Enter += new System.EventHandler(this.New6YRYear_Enter);
            this.New6YRYear.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            this.New6YRYear.Leave += new System.EventHandler(this.New6YRYear_Leave);
            // 
            // NewHTYear
            // 
            this.NewHTYear.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.NewHTYear.Location = new System.Drawing.Point(390, 196);
            this.NewHTYear.Name = "NewHTYear";
            this.NewHTYear.Size = new System.Drawing.Size(100, 20);
            this.NewHTYear.TabIndex = 8;
            this.NewHTYear.Enter += new System.EventHandler(this.NewHTYear_Enter);
            this.NewHTYear.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            this.NewHTYear.Leave += new System.EventHandler(this.NewHTYear_Leave);
            // 
            // NewExtraLabor
            // 
            this.NewExtraLabor.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.NewExtraLabor.Location = new System.Drawing.Point(270, 196);
            this.NewExtraLabor.Name = "NewExtraLabor";
            this.NewExtraLabor.Size = new System.Drawing.Size(100, 20);
            this.NewExtraLabor.TabIndex = 7;
            this.NewExtraLabor.TextChanged += new System.EventHandler(this.NewExtraLabor_TextChanged);
            this.NewExtraLabor.Enter += new System.EventHandler(this.NewExtraLabor_Enter);
            this.NewExtraLabor.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // NewExtraParts
            // 
            this.NewExtraParts.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.NewExtraParts.Location = new System.Drawing.Point(140, 196);
            this.NewExtraParts.Name = "NewExtraParts";
            this.NewExtraParts.Size = new System.Drawing.Size(100, 20);
            this.NewExtraParts.TabIndex = 6;
            this.NewExtraParts.TextChanged += new System.EventHandler(this.NewExtraParts_TextChanged);
            this.NewExtraParts.Enter += new System.EventHandler(this.NewExtraParts_Enter);
            this.NewExtraParts.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // NewChemical
            // 
            this.NewChemical.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.NewChemical.Location = new System.Drawing.Point(630, 136);
            this.NewChemical.Name = "NewChemical";
            this.NewChemical.Size = new System.Drawing.Size(100, 20);
            this.NewChemical.TabIndex = 5;
            this.NewChemical.TextChanged += new System.EventHandler(this.NewChemical_TextChanged);
            this.NewChemical.Enter += new System.EventHandler(this.NewChemical_Enter);
            this.NewChemical.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // NewSize
            // 
            this.NewSize.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.NewSize.Location = new System.Drawing.Point(510, 137);
            this.NewSize.Name = "NewSize";
            this.NewSize.Size = new System.Drawing.Size(100, 20);
            this.NewSize.TabIndex = 4;
            this.NewSize.TextChanged += new System.EventHandler(this.NewSize_TextChanged);
            this.NewSize.Enter += new System.EventHandler(this.NewSize_Enter);
            this.NewSize.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // NewOR2
            // 
            this.NewOR2.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.NewOR2.Location = new System.Drawing.Point(390, 136);
            this.NewOR2.Name = "NewOR2";
            this.NewOR2.Size = new System.Drawing.Size(100, 20);
            this.NewOR2.TabIndex = 3;
            this.NewOR2.TextChanged += new System.EventHandler(this.NewOR2_TextChanged);
            this.NewOR2.Enter += new System.EventHandler(this.NewOR2_Enter);
            this.NewOR2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // NewOR1
            // 
            this.NewOR1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.NewOR1.Location = new System.Drawing.Point(270, 136);
            this.NewOR1.Name = "NewOR1";
            this.NewOR1.Size = new System.Drawing.Size(100, 20);
            this.NewOR1.TabIndex = 2;
            this.NewOR1.TextChanged += new System.EventHandler(this.NewOR1_TextChanged);
            this.NewOR1.Enter += new System.EventHandler(this.NewOR1_Enter);
            this.NewOR1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // NewVS
            // 
            this.NewVS.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.NewVS.Location = new System.Drawing.Point(140, 136);
            this.NewVS.Name = "NewVS";
            this.NewVS.Size = new System.Drawing.Size(100, 20);
            this.NewVS.TabIndex = 1;
            this.NewVS.TextChanged += new System.EventHandler(this.NewVS_TextChanged);
            this.NewVS.Enter += new System.EventHandler(this.NewVS_Enter);
            this.NewVS.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // NewModelNum
            // 
            this.NewModelNum.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.NewModelNum.Location = new System.Drawing.Point(10, 137);
            this.NewModelNum.Name = "NewModelNum";
            this.NewModelNum.Size = new System.Drawing.Size(100, 20);
            this.NewModelNum.TabIndex = 0;
            this.NewModelNum.TextChanged += new System.EventHandler(this.NewModelNum_TextChanged);
            this.NewModelNum.Enter += new System.EventHandler(this.NewModelNum_Enter);
            this.NewModelNum.KeyDown += new System.Windows.Forms.KeyEventHandler(this.Control_KeyUp);
            // 
            // NewTemplate
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(744, 450);
            this.Controls.Add(this.panel1);
            this.Name = "NewTemplate";
            this.Text = "New Template";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox NewModelNum;
        private System.Windows.Forms.TextBox New6YRYear;
        private System.Windows.Forms.TextBox NewHTYear;
        private System.Windows.Forms.TextBox NewExtraLabor;
        private System.Windows.Forms.TextBox NewExtraParts;
        private System.Windows.Forms.TextBox NewChemical;
        private System.Windows.Forms.TextBox NewSize;
        private System.Windows.Forms.TextBox NewOR2;
        private System.Windows.Forms.TextBox NewOR1;
        private System.Windows.Forms.TextBox NewVS;
        private System.Windows.Forms.Label NewLabel;
        private System.Windows.Forms.Label NewCollar;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label Title;
        private System.Windows.Forms.RadioButton NewCTNoB;
        private System.Windows.Forms.RadioButton NewCTYesB;
        private System.Windows.Forms.RadioButton NewCollarNoB;
        private System.Windows.Forms.RadioButton NewCollarYesB;
        private Button SubmissionB;
        private GroupBox groupBox2;
        private GroupBox groupBox1;
        private CheckBox multiSizeBox;
    }
}