namespace Resx2Xls
{
    partial class Resx2XlsForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.resxDirectoryDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.wizardControl1 = new WizardBase.WizardControl();
            this.startStep1 = new WizardBase.StartStep();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.radioButtonBuildXls = new System.Windows.Forms.RadioButton();
            this.radioButtonCreateXls = new System.Windows.Forms.RadioButton();
            this.intermediateStepProject = new WizardBase.IntermediateStep();
            this.labelFolder = new System.Windows.Forms.Label();
            this.textBoxResxFolder = new System.Windows.Forms.TextBox();
            this.buttonBrowse = new System.Windows.Forms.Button();
            this.intermediateStepCultures = new WizardBase.IntermediateStep();
            this.buttonRemove = new System.Windows.Forms.Button();
            this.buttonRemoveAll = new System.Windows.Forms.Button();
            this.buttonAdd = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.buttonAddAll = new System.Windows.Forms.Button();
            this.listBoxCultures = new System.Windows.Forms.ListBox();
            this.listBoxSelected = new System.Windows.Forms.ListBox();
            this.intermediateStepXlsSelect = new WizardBase.IntermediateStep();
            this.labelXlsFile = new System.Windows.Forms.Label();
            this.textBoxXlsDirectory = new System.Windows.Forms.TextBox();
            this.buttonBrowseXlsDirectory = new System.Windows.Forms.Button();
            this.finishStep1 = new WizardBase.FinishStep();
            this.label6 = new System.Windows.Forms.Label();
            this.textBoxSummary = new System.Windows.Forms.TextBox();
            this.xlsDirectoryDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.startStep1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.intermediateStepProject.SuspendLayout();
            this.intermediateStepCultures.SuspendLayout();
            this.intermediateStepXlsSelect.SuspendLayout();
            this.finishStep1.SuspendLayout();
            this.SuspendLayout();
            // 
            // wizardControl1
            // 
            this.wizardControl1.BackButtonEnabled = true;
            this.wizardControl1.BackButtonVisible = true;
            this.wizardControl1.CancelButtonEnabled = true;
            this.wizardControl1.CancelButtonVisible = true;
            this.wizardControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.wizardControl1.HelpButtonEnabled = true;
            this.wizardControl1.HelpButtonVisible = false;
            this.wizardControl1.Location = new System.Drawing.Point(0, 0);
            this.wizardControl1.Name = "wizardControl1";
            this.wizardControl1.NextButtonEnabled = true;
            this.wizardControl1.NextButtonVisible = true;
            this.wizardControl1.Size = new System.Drawing.Size(704, 466);
            this.wizardControl1.WizardSteps.Add(this.startStep1);
            this.wizardControl1.WizardSteps.Add(this.intermediateStepProject);
            this.wizardControl1.WizardSteps.Add(this.intermediateStepCultures);
            this.wizardControl1.WizardSteps.Add(this.intermediateStepXlsSelect);
            this.wizardControl1.WizardSteps.Add(this.finishStep1);
            this.wizardControl1.BackButtonClick += new WizardBase.WizardClickEventHandler(this.wizardControl1_BackButtonClick);
            this.wizardControl1.CancelButtonClick += new System.EventHandler(this.wizardControl1_CancelButtonClick);
            this.wizardControl1.FinishButtonClick += new System.EventHandler(this.wizardControl1_FinishButtonClick);
            this.wizardControl1.NextButtonClick += new WizardBase.WizardNextButtonClickEventHandler(this.wizardControl1_NextButtonClick);
            // 
            // startStep1
            // 
            this.startStep1.BindingImage = global::Resx2Xls.Properties.Resources.leftbar;
            this.startStep1.Controls.Add(this.groupBox1);
            this.startStep1.Icon = global::Resx2Xls.Properties.Resources.icon;
            this.startStep1.Name = "startStep1";
            this.startStep1.Subtitle = "This wizard helps you to localize your .Net Project";
            this.startStep1.SubtitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.startStep1.Title = "Welcome to the Resx to Xls Wizard.";
            this.startStep1.TitleFont = new System.Drawing.Font("Verdana", 12F, System.Drawing.FontStyle.Bold);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.radioButtonBuildXls);
            this.groupBox1.Controls.Add(this.radioButtonCreateXls);
            this.groupBox1.Location = new System.Drawing.Point(198, 93);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(373, 100);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Options";
            // 
            // radioButtonBuildXls
            // 
            this.radioButtonBuildXls.AutoSize = true;
            this.radioButtonBuildXls.Location = new System.Drawing.Point(45, 52);
            this.radioButtonBuildXls.Name = "radioButtonBuildXls";
            this.radioButtonBuildXls.Size = new System.Drawing.Size(161, 17);
            this.radioButtonBuildXls.TabIndex = 1;
            this.radioButtonBuildXls.Text = "Import transaltions from excel";
            this.radioButtonBuildXls.UseVisualStyleBackColor = true;
            // 
            // radioButtonCreateXls
            // 
            this.radioButtonCreateXls.AutoSize = true;
            this.radioButtonCreateXls.Checked = true;
            this.radioButtonCreateXls.Location = new System.Drawing.Point(45, 29);
            this.radioButtonCreateXls.Name = "radioButtonCreateXls";
            this.radioButtonCreateXls.Size = new System.Drawing.Size(151, 17);
            this.radioButtonCreateXls.TabIndex = 0;
            this.radioButtonCreateXls.TabStop = true;
            this.radioButtonCreateXls.Text = "Export translations to excel\r\n";
            this.radioButtonCreateXls.UseVisualStyleBackColor = true;
            // 
            // intermediateStepProject
            // 
            this.intermediateStepProject.BindingImage = global::Resx2Xls.Properties.Resources.topbar;
            this.intermediateStepProject.Controls.Add(this.labelFolder);
            this.intermediateStepProject.Controls.Add(this.textBoxResxFolder);
            this.intermediateStepProject.Controls.Add(this.buttonBrowse);
            this.intermediateStepProject.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.intermediateStepProject.Name = "intermediateStepProject";
            this.intermediateStepProject.Subtitle = "Browse the resx folder of your .Net Project..";
            this.intermediateStepProject.SubtitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.intermediateStepProject.Title = "Select your .Net Project.";
            this.intermediateStepProject.TitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
            // 
            // labelFolder
            // 
            this.labelFolder.AutoSize = true;
            this.labelFolder.ForeColor = System.Drawing.SystemColors.ControlText;
            this.labelFolder.Location = new System.Drawing.Point(20, 120);
            this.labelFolder.Name = "labelFolder";
            this.labelFolder.Size = new System.Drawing.Size(250, 13);
            this.labelFolder.TabIndex = 10;
            this.labelFolder.Text = "Project Resx Folder (that contains neutral resx files):";
            // 
            // textBoxResxFolder
            // 
            this.textBoxResxFolder.Location = new System.Drawing.Point(23, 136);
            this.textBoxResxFolder.Name = "textBoxResxFolder";
            this.textBoxResxFolder.Size = new System.Drawing.Size(438, 20);
            this.textBoxResxFolder.TabIndex = 9;
            this.textBoxResxFolder.TextChanged += new System.EventHandler(this.textBoxResxFolder_TextChanged);
            // 
            // buttonBrowse
            // 
            this.buttonBrowse.Location = new System.Drawing.Point(467, 136);
            this.buttonBrowse.Name = "buttonBrowse";
            this.buttonBrowse.Size = new System.Drawing.Size(75, 23);
            this.buttonBrowse.TabIndex = 11;
            this.buttonBrowse.Text = "Browse";
            this.buttonBrowse.UseVisualStyleBackColor = true;
            this.buttonBrowse.Click += new System.EventHandler(this.buttonBrowse_Click);
            // 
            // intermediateStepCultures
            // 
            this.intermediateStepCultures.BindingImage = global::Resx2Xls.Properties.Resources.topbar;
            this.intermediateStepCultures.Controls.Add(this.buttonRemove);
            this.intermediateStepCultures.Controls.Add(this.buttonRemoveAll);
            this.intermediateStepCultures.Controls.Add(this.buttonAdd);
            this.intermediateStepCultures.Controls.Add(this.label5);
            this.intermediateStepCultures.Controls.Add(this.label4);
            this.intermediateStepCultures.Controls.Add(this.label3);
            this.intermediateStepCultures.Controls.Add(this.label1);
            this.intermediateStepCultures.Controls.Add(this.buttonAddAll);
            this.intermediateStepCultures.Controls.Add(this.listBoxCultures);
            this.intermediateStepCultures.Controls.Add(this.listBoxSelected);
            this.intermediateStepCultures.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.intermediateStepCultures.Name = "intermediateStepCultures";
            this.intermediateStepCultures.Subtitle = "This step creates a new xls file per culture that contains all your resource keys" +
    ".";
            this.intermediateStepCultures.SubtitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.intermediateStepCultures.Title = "Select the Cultures that you want include in the project.";
            this.intermediateStepCultures.TitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
            // 
            // buttonRemove
            // 
            this.buttonRemove.ForeColor = System.Drawing.SystemColors.ControlText;
            this.buttonRemove.Location = new System.Drawing.Point(222, 158);
            this.buttonRemove.Name = "buttonRemove";
            this.buttonRemove.Size = new System.Drawing.Size(52, 23);
            this.buttonRemove.TabIndex = 13;
            this.buttonRemove.Text = "<";
            this.buttonRemove.UseVisualStyleBackColor = true;
            this.buttonRemove.Click += new System.EventHandler(this.buttonRemove_Click);
            // 
            // buttonRemoveAll
            // 
            this.buttonRemoveAll.ForeColor = System.Drawing.SystemColors.ControlText;
            this.buttonRemoveAll.Location = new System.Drawing.Point(222, 187);
            this.buttonRemoveAll.Name = "buttonRemoveAll";
            this.buttonRemoveAll.Size = new System.Drawing.Size(52, 23);
            this.buttonRemoveAll.TabIndex = 12;
            this.buttonRemoveAll.Text = "<<";
            this.buttonRemoveAll.UseVisualStyleBackColor = true;
            this.buttonRemoveAll.Click += new System.EventHandler(this.buttonRemoveAll_Click);
            // 
            // buttonAdd
            // 
            this.buttonAdd.ForeColor = System.Drawing.SystemColors.ControlText;
            this.buttonAdd.Location = new System.Drawing.Point(222, 129);
            this.buttonAdd.Name = "buttonAdd";
            this.buttonAdd.Size = new System.Drawing.Size(52, 23);
            this.buttonAdd.TabIndex = 11;
            this.buttonAdd.Text = ">";
            this.buttonAdd.UseVisualStyleBackColor = true;
            this.buttonAdd.Click += new System.EventHandler(this.buttonAdd_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label5.Location = new System.Drawing.Point(277, 355);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(160, 13);
            this.label5.TabIndex = 10;
            this.label5.Text = "Double click to remove a culture";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label4.Location = new System.Drawing.Point(64, 355);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(143, 13);
            this.label4.TabIndex = 9;
            this.label4.Text = "Double click to add a culture";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label3.Location = new System.Drawing.Point(277, 84);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(93, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Selected Cultures:";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.ForeColor = System.Drawing.SystemColors.ControlText;
            this.label1.Location = new System.Drawing.Point(64, 84);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(94, 13);
            this.label1.TabIndex = 5;
            this.label1.Text = "Available Cultures:";
            // 
            // buttonAddAll
            // 
            this.buttonAddAll.ForeColor = System.Drawing.SystemColors.ControlText;
            this.buttonAddAll.Location = new System.Drawing.Point(222, 100);
            this.buttonAddAll.Name = "buttonAddAll";
            this.buttonAddAll.Size = new System.Drawing.Size(52, 23);
            this.buttonAddAll.TabIndex = 7;
            this.buttonAddAll.Text = ">>";
            this.buttonAddAll.UseVisualStyleBackColor = true;
            this.buttonAddAll.Click += new System.EventHandler(this.buttonAddAll_Click);
            // 
            // listBoxCultures
            // 
            this.listBoxCultures.FormattingEnabled = true;
            this.listBoxCultures.Location = new System.Drawing.Point(67, 100);
            this.listBoxCultures.Name = "listBoxCultures";
            this.listBoxCultures.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.listBoxCultures.Size = new System.Drawing.Size(149, 251);
            this.listBoxCultures.TabIndex = 4;
            this.listBoxCultures.SelectedIndexChanged += new System.EventHandler(this.listBoxCultures_SelectedIndexChanged);
            this.listBoxCultures.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.listBoxCultures_MouseDoubleClick);
            // 
            // listBoxSelected
            // 
            this.listBoxSelected.FormattingEnabled = true;
            this.listBoxSelected.Location = new System.Drawing.Point(280, 100);
            this.listBoxSelected.Name = "listBoxSelected";
            this.listBoxSelected.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.listBoxSelected.Size = new System.Drawing.Size(149, 251);
            this.listBoxSelected.TabIndex = 6;
            this.listBoxSelected.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.listBoxSelected_MouseDoubleClick);
            // 
            // intermediateStepXlsSelect
            // 
            this.intermediateStepXlsSelect.BindingImage = global::Resx2Xls.Properties.Resources.topbar;
            this.intermediateStepXlsSelect.Controls.Add(this.labelXlsFile);
            this.intermediateStepXlsSelect.Controls.Add(this.textBoxXlsDirectory);
            this.intermediateStepXlsSelect.Controls.Add(this.buttonBrowseXlsDirectory);
            this.intermediateStepXlsSelect.ForeColor = System.Drawing.SystemColors.HighlightText;
            this.intermediateStepXlsSelect.Name = "intermediateStepXlsSelect";
            this.intermediateStepXlsSelect.Subtitle = "Give a valid document directory that will contain localization info.";
            this.intermediateStepXlsSelect.SubtitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.intermediateStepXlsSelect.Title = "Select your translations directory.";
            this.intermediateStepXlsSelect.TitleFont = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);
            // 
            // labelXlsFile
            // 
            this.labelXlsFile.AutoSize = true;
            this.labelXlsFile.ForeColor = System.Drawing.SystemColors.ControlText;
            this.labelXlsFile.Location = new System.Drawing.Point(32, 94);
            this.labelXlsFile.Name = "labelXlsFile";
            this.labelXlsFile.Size = new System.Drawing.Size(135, 13);
            this.labelXlsFile.TabIndex = 2;
            this.labelXlsFile.Text = "Excel translations directory:";
            // 
            // textBoxXlsDirectory
            // 
            this.textBoxXlsDirectory.Location = new System.Drawing.Point(35, 110);
            this.textBoxXlsDirectory.Name = "textBoxXlsDirectory";
            this.textBoxXlsDirectory.Size = new System.Drawing.Size(385, 20);
            this.textBoxXlsDirectory.TabIndex = 0;
            this.textBoxXlsDirectory.TextChanged += new System.EventHandler(this.textBoxXlsDirectory_TextChanged);
            // 
            // buttonBrowseXlsDirectory
            // 
            this.buttonBrowseXlsDirectory.ForeColor = System.Drawing.SystemColors.ControlText;
            this.buttonBrowseXlsDirectory.Location = new System.Drawing.Point(426, 108);
            this.buttonBrowseXlsDirectory.Name = "buttonBrowseXlsDirectory";
            this.buttonBrowseXlsDirectory.Size = new System.Drawing.Size(75, 23);
            this.buttonBrowseXlsDirectory.TabIndex = 1;
            this.buttonBrowseXlsDirectory.Text = "Browse";
            this.buttonBrowseXlsDirectory.UseVisualStyleBackColor = true;
            this.buttonBrowseXlsDirectory.Click += new System.EventHandler(this.buttonBrowseXlsDirectory_Click);
            // 
            // finishStep1
            // 
            this.finishStep1.BackgroundImage = global::Resx2Xls.Properties.Resources.finishbar;
            this.finishStep1.Controls.Add(this.label6);
            this.finishStep1.Controls.Add(this.textBoxSummary);
            this.finishStep1.Name = "finishStep1";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(24, 88);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(53, 13);
            this.label6.TabIndex = 1;
            this.label6.Text = "Summary:";
            // 
            // textBoxSummary
            // 
            this.textBoxSummary.Location = new System.Drawing.Point(27, 103);
            this.textBoxSummary.Multiline = true;
            this.textBoxSummary.Name = "textBoxSummary";
            this.textBoxSummary.ReadOnly = true;
            this.textBoxSummary.Size = new System.Drawing.Size(646, 255);
            this.textBoxSummary.TabIndex = 0;
            // 
            // Resx2XlsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(704, 466);
            this.Controls.Add(this.wizardControl1);
            this.Name = "Resx2XlsForm";
            this.Text = "Resx To Xls";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Resx2XlsForm_FormClosing);
            this.startStep1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.intermediateStepProject.ResumeLayout(false);
            this.intermediateStepProject.PerformLayout();
            this.intermediateStepCultures.ResumeLayout(false);
            this.intermediateStepCultures.PerformLayout();
            this.intermediateStepXlsSelect.ResumeLayout(false);
            this.intermediateStepXlsSelect.PerformLayout();
            this.finishStep1.ResumeLayout(false);
            this.finishStep1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog resxDirectoryDialog;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox listBoxCultures;
        private System.Windows.Forms.Button buttonAddAll;
        private System.Windows.Forms.ListBox listBoxSelected;
        private System.Windows.Forms.Label labelXlsFile;
        private System.Windows.Forms.Button buttonBrowseXlsDirectory;
        private System.Windows.Forms.TextBox textBoxXlsDirectory;
        private System.Windows.Forms.Button buttonBrowse;
        private System.Windows.Forms.TextBox textBoxResxFolder;
        private System.Windows.Forms.Label labelFolder;
        private WizardBase.WizardControl wizardControl1;
        private WizardBase.StartStep startStep1;
        private WizardBase.IntermediateStep intermediateStepProject;
        private WizardBase.IntermediateStep intermediateStepCultures;
        private WizardBase.FinishStep finishStep1;
        private WizardBase.IntermediateStep intermediateStepXlsSelect;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton radioButtonBuildXls;
        private System.Windows.Forms.RadioButton radioButtonCreateXls;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxSummary;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.FolderBrowserDialog xlsDirectoryDialog;
        private System.Windows.Forms.Button buttonAdd;
        private System.Windows.Forms.Button buttonRemove;
        private System.Windows.Forms.Button buttonRemoveAll;
    }
}

