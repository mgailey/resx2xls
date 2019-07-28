using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using Resx2Xls.Properties;
using WizardBase;

namespace Resx2Xls
{
    public partial class Resx2XlsForm : Form
    {
        private readonly string summary1;
        private readonly string summary2;
        private ResxToXlsOperation operation;
        private readonly Dictionary<int, NavLink> exportNav = new Dictionary<int, NavLink>
                                                             {
                                                                 {0, new NavLink{Next = 1}},
                                                                 {1, new NavLink{Next = 2, Previous = 0}},
                                                                 {2, new NavLink{Next = 3, Previous = 1}},
                                                                 {3, new NavLink{Next = 4, Previous = 2}},
                                                                 {4, new NavLink{Previous = 3}},
                                                             };
        private readonly Dictionary<int, NavLink> importNav = new Dictionary<int, NavLink>
                                                             {
                                                                 {0, new NavLink{Next = 3}},
                                                                 {1, new NavLink{Next = 4, Previous = 2}},
                                                                 {2, new NavLink{Next = 1, Previous = 3}},
                                                                 {3, new NavLink{Next = 2, Previous = 0}},
                                                                 {4, new NavLink{Previous = 1}},
                                                             };

        public Resx2XlsForm()
        {
            InitializeComponent();

            textBoxResxFolder.Text = Settings.Default.ResxFolderPath;
            textBoxXlsDirectory.Text = Settings.Default.XlsFolderPath;

            SetOperationState();
            radioButtonCreateXls.CheckedChanged += radioButton_CheckedChanged;
            radioButtonBuildXls.CheckedChanged += radioButton_CheckedChanged;

            summary1 = "Operation:\r\nCreate a new Excel document ready for localization.";
            summary2 = "Operation:\r\nBuild your localized Resource files from a filled Excel Document.";

            textBoxSummary.Text = summary1;
        }

        private void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            radioButtonCreateXls.CheckedChanged -= radioButton_CheckedChanged;
            radioButtonBuildXls.CheckedChanged -= radioButton_CheckedChanged;

            SetOperationState();

            radioButtonCreateXls.CheckedChanged += radioButton_CheckedChanged;
            radioButtonBuildXls.CheckedChanged += radioButton_CheckedChanged;
        }

        private void SetOperationState()
        {
            if (radioButtonCreateXls.Checked)
            {
                operation = ResxToXlsOperation.Export;
                textBoxSummary.Text = summary1;
                listBoxCultures.AddCulturesFrom(textBoxResxFolder.Text);
            }
            if (radioButtonBuildXls.Checked)
            {
                operation = ResxToXlsOperation.ImportFile;
                textBoxSummary.Text = summary2;
                listBoxCultures.AddCulturesFrom(textBoxXlsDirectory.Text, "*.xlsx");
            }
        }

        private void AddSelectedCultures()
        {
            foreach (var ci in listBoxCultures.SelectedItems.Cast<CultureInfo>().Where(ci => listBoxSelected.Items.IndexOf(ci) == -1))
            {
                listBoxSelected.Items.Add(ci);
            }
        }

        private void SaveCultures()
        {
            Settings.Default.CultureList = string.Join(";", listBoxSelected.Items.Cast<CultureInfo>().Select(ci => ci.Name).ToArray());
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            if (resxDirectoryDialog.ShowDialog() != DialogResult.OK) return;
                
            textBoxResxFolder.Text = resxDirectoryDialog.SelectedPath;
            listBoxCultures.AddCulturesFrom(textBoxResxFolder.Text);
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            AddSelectedCultures();
        }

        private void listBoxCultures_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            AddSelectedCultures();
        }

        private void buttonBrowseXlsDirectory_Click(object sender, EventArgs e)
        {
            if (xlsDirectoryDialog.ShowDialog() != DialogResult.OK) return;

            textBoxXlsDirectory.Text = xlsDirectoryDialog.SelectedPath;
            listBoxCultures.AddCulturesFrom(textBoxXlsDirectory.Text, "*.xlsx");
        }

        private void listBoxSelected_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            RemoveSelectedCultures();
        }

        private void RemoveSelectedCultures()
        {
            if (listBoxSelected.SelectedItems.Count <= 0) return;
            var selectedItems = new object[listBoxSelected.SelectedItems.Count];
            listBoxSelected.SelectedItems.CopyTo(selectedItems, 0);
            foreach (var selectedItem in selectedItems)
            {
                listBoxSelected.Items.Remove(selectedItem);
            }
        }

        private void Resx2XlsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveCultures();
            Settings.Default.Save();
        }

        private void textBoxResxFolder_TextChanged(object sender, EventArgs e)
        {
            Settings.Default.ResxFolderPath = textBoxResxFolder.Text;
        }

        private void textBoxXlsDirectory_TextChanged(object sender, EventArgs e)
        {
            Settings.Default.XlsFolderPath = textBoxXlsDirectory.Text;
        }

        private void FinishWizard()
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                var cultures = (from object selected in listBoxSelected.Items select ((CultureInfo)selected).Name).ToArray();
                switch (operation)
                {
                    case ResxToXlsOperation.Export:
                        FinishExport(cultures);
                        break;
                    case ResxToXlsOperation.ImportFile:
                        FinishXlsImport(cultures);
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                Cursor = Cursors.Default;  
            }

            Close();
        }

        private void FinishXlsImport(string[] cultures)
        {
            if (string.IsNullOrEmpty(textBoxXlsDirectory.Text))
            {
                MessageBox.Show("You must select the Excel directory to import from", "Update",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information);
                wizardControl1.CurrentStepIndex = 4;
                return;
            }

            var import = new ResxImporter(new DirectoryInfo(textBoxResxFolder.Text), new DirectoryInfo(textBoxXlsDirectory.Text));
            import.Import(cultures);

            MessageBox.Show("Resources imported.", "Import", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
        }

        private void FinishExport(string[] cultures)
        {
            if (string.IsNullOrEmpty(textBoxResxFolder.Text))
            {
                MessageBox.Show("You must select a the .Net Project resx wich contains your updated resx files", "Update",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                wizardControl1.CurrentStepIndex = 1;
                return;
            }

            if (string.IsNullOrEmpty(textBoxXlsDirectory.Text))
            {
                MessageBox.Show("You must select the directory for your xls files", "Update",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                wizardControl1.CurrentStepIndex = 5;
                return;
            }

            var export = new ResxExporter(new DirectoryInfo(textBoxResxFolder.Text), new DirectoryInfo(textBoxXlsDirectory.Text));
            export.Export(cultures);
            MessageBox.Show("Localized Resources exported.", "Export", MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
        }

        private void wizardControl1_NextButtonClick(WizardControl sender, WizardNextButtonClickEventArgs args)
        {
            var index = wizardControl1.CurrentStepIndex;
            var current = GetNavLink(index);
            wizardControl1.CurrentStepIndex = current.Next - 1; //position index on step before desired index
        }

        private void wizardControl1_BackButtonClick(WizardControl sender, WizardClickEventArgs args)
        {
            var index = wizardControl1.CurrentStepIndex;
            var current = GetNavLink(index);
            wizardControl1.CurrentStepIndex = current.Previous + 1; //position index on step after desired index
        }

        private NavLink GetNavLink(int index)
        {
            var current = new NavLink();
            switch (operation)
            {
                case ResxToXlsOperation.Export:
                    current = exportNav[index];
                    break;
                case ResxToXlsOperation.ImportFile:
                    current = importNav[index];
                    break;
            }
            return current;
        }

        private void wizardControl1_FinishButtonClick(object sender, EventArgs e)
        {
            FinishWizard();
        }

        private void wizardControl1_CancelButtonClick(object sender, EventArgs e)
        {
            Close();
        }

        private enum ResxToXlsOperation
        {
            Export,
            ImportFile
        };

        private struct NavLink
        {
            public int Previous { get; set; }
            public int Next { get; set; }
        }

        private void buttonAddAll_Click(object sender, EventArgs e)
        {
            AddAllCultures();
        }

        private void AddAllCultures()
        {
            RemoveAllSelectedCultures();
            foreach (var ci in listBoxCultures.Items.Cast<CultureInfo>())
            {
                listBoxSelected.Items.Add(ci);
            }
        }

        private void RemoveAllSelectedCultures()
        {
            listBoxSelected.Items.Clear();
        }

        private void buttonRemoveAll_Click(object sender, EventArgs e)
        {
            RemoveAllSelectedCultures();
        }

        private void buttonRemove_Click(object sender, EventArgs e)
        {
            RemoveSelectedCultures();
        }

        private void listBoxCultures_SelectedIndexChanged(object sender, EventArgs e)
        {
            AddSelectedCultures();
        }
    }

}