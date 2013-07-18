using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Resx2Xls.Properties;
using WizardBase;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace Resx2Xls
{
    public partial class Resx2XlsForm : Form
    {
        private readonly string summary1;
        private readonly string summary2;
        private readonly string summary3;
        private readonly object objOpt = Missing.Value;
        private ResxToXlsOperation operation;


        public Resx2XlsForm()
        {
            InitializeComponent();

            textBoxFolder.Text = Settings.Default.FolderPath;
            textBoxExclude.Text = Settings.Default.ExcludeList;

            FillCultures();
            AddExistingCultures();

            radioButtonCreateXls.CheckedChanged += radioButton_CheckedChanged;
            radioButtonBuildXls.CheckedChanged += radioButton_CheckedChanged;
            radioButtonUpdateXls.CheckedChanged += radioButton_CheckedChanged;

            summary1 = "Operation:\r\nCreate a new Excel document ready for localization.";
            summary2 = "Operation:\r\nBuild your localized Resource files from a Filled Excel Document.";
            summary3 = "Operation:\r\nUpdate your Excel document with your Project Resource changes.";

            textBoxSummary.Text = summary1;
        }

        private void radioButton_CheckedChanged(object sender, EventArgs e)
        {
            radioButtonCreateXls.CheckedChanged -= radioButton_CheckedChanged;
            radioButtonBuildXls.CheckedChanged -= radioButton_CheckedChanged;
            radioButtonUpdateXls.CheckedChanged -= radioButton_CheckedChanged;

            if (radioButtonCreateXls.Checked)
            {
                operation = ResxToXlsOperation.Export;
                textBoxSummary.Text = summary1;
            }
            if (radioButtonBuildXls.Checked)
            {
                operation = ResxToXlsOperation.ImportFile;
                textBoxSummary.Text = summary2;
            }
            if (radioButtonUpdateXls.Checked)
            {
                operation = ResxToXlsOperation.ImportDirectory;
                textBoxSummary.Text = summary3;
            }

            if (((RadioButton) sender).Checked)
            {
                if ((sender) == radioButtonCreateXls)
                {
                    radioButtonBuildXls.Checked = false;
                    radioButtonUpdateXls.Checked = false;
                }

                if ((sender) == radioButtonBuildXls)
                {
                    radioButtonCreateXls.Checked = false;
                    radioButtonUpdateXls.Checked = false;
                }

                if ((sender) == radioButtonUpdateXls)
                {
                    radioButtonCreateXls.Checked = false;
                    radioButtonBuildXls.Checked = false;
                }
            }
            radioButtonCreateXls.CheckedChanged += radioButton_CheckedChanged;
            radioButtonBuildXls.CheckedChanged += radioButton_CheckedChanged;
            radioButtonUpdateXls.CheckedChanged += radioButton_CheckedChanged;
        }

        public void ResxToXls(string sourceDirectory, string destDirectory, string[] cultures, string[] excludeList)
        {
            if (!Directory.Exists(sourceDirectory))
                return;
            var rd = new Resx();

            var file = FindRootResxFile(sourceDirectory);
            ReadNeutralResx(file, rd, excludeList);
            foreach (var culture in cultures)
            {
                var cultureFile = FindCultureResxFile(sourceDirectory, culture);
                if (cultureFile != null)
                {
                    AppendCulture(cultureFile, culture, rd, excludeList); 
                }
                else
                {
                    throw new Exception("File not found");
                }
                var destFile = new FileInfo(Path.Combine(destDirectory, string.Format("{0}.xlsx", JustStem(cultureFile.Name))));
                DataSetToXls(rd, destFile, culture);   
            }

            ShowXls(destDirectory);
        }

        private void XlsToResx(string xlsFile)
        {
            if (!File.Exists(xlsFile))
                return;

            var path = new FileInfo(xlsFile).DirectoryName;

            var app = new Application();
            var wb = app.Workbooks.Open(xlsFile,
                                             0, false, 5, "", "", false, XlPlatform.xlWindows, "",
                                             true, false, 0, true, false, false);

            var sheets = wb.Worksheets;

            var sheet = (Worksheet) sheets.Item[1];

            var hasLanguage = true;
            var col = 5;

            while (hasLanguage)
            {
                var val = ((Range) sheet.Cells[2, col]).Text;

                if (val is string)
                {
                    if (!String.IsNullOrEmpty((string) val))
                    {
                        var cult = (string) val;

                        var pathCulture = path + @"\" + cult;

                        if (!Directory.Exists(pathCulture))
                            Directory.CreateDirectory(pathCulture);


                        var row = 3;

                        var readrow = true;

                        while (readrow)
                        {
                            var range = sheet.Cells[row, 2] as Range;
                            if (range == null) continue;

                            var fileDest = range.Text.ToString();

                            if (String.IsNullOrEmpty(fileDest))
                                break;

                            var f = pathCulture + @"\" + JustStem(fileDest) + "." + cult + ".resx";

                            var rw = new ResXResourceWriter(f);

                            while (readrow)
                            {
                                var range1 = sheet.Cells[row, 3] as Range;
                                if (range1 == null) continue;

                                var key = range1.Text.ToString();
                                var range2 = sheet.Cells[row, col] as Range;
                                if (range2 == null) continue;

                                object data = range2.Text.ToString();

                                if (!String.IsNullOrEmpty(key))
                                {
                                    var text = (string) data;

                                    text = text.Replace("\\r", "\r");
                                    text = text.Replace("\\n", "\n");

                                    rw.AddResource(new ResXDataNode(key, text));

                                    row++;

                                    var file = range.Text.ToString();

                                    if (file != fileDest)
                                        break;
                                }
                                else
                                {
                                    readrow = false;
                                }
                            }

                            rw.Close();
                        }
                    }
                    else
                        hasLanguage = false;
                }
                else
                    hasLanguage = false;

                col++;
            }
        }

        private FileInfo FindRootResxFile(string path)
        {
            var files = Directory.GetFiles(path, "*.resx", SearchOption.TopDirectoryOnly);
            foreach (var f in files)
            {
                CultureInfo cult;
                if (!IsResxCultureSpecific(f, out cult))
                {
                    return new FileInfo(f);
                }
            }
            return null;
        }

        private FileInfo FindCultureResxFile(string path, string culture)
        {
            var files = Directory.GetFiles(path, string.Format("*.{0}.resx", culture), SearchOption.TopDirectoryOnly);
            foreach (var f in files)
            {
                CultureInfo cult;
                if (IsResxCultureSpecific(f, out cult) && cult.Name == culture) return new FileInfo(f);
            }
            return null;
        }

        private bool IsResxCultureSpecific(string path, out CultureInfo culture)
        {
            culture = null;
            var fi = new FileInfo(path);

            //Remove the extension and return the string	
            var fname = JustStem(fi.Name);

            var cult = String.Empty;
            if (fname.IndexOf(".") != -1)
                cult = fname.Substring(fname.LastIndexOf('.') + 1);

            if (cult == String.Empty)
                return false;

            try
            {
                culture = new CultureInfo(cult);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private void ReadNeutralResx(FileSystemInfo neutralFile, Resx rd, string[] excludeList)
        {
            var reader = new ResXResourceReader(neutralFile.FullName);

            try
            {
                foreach (DictionaryEntry de in reader)
                {
                    if (!(de.Value is string)) continue;

                    var key = (string) de.Key;
                    var exclude = excludeList.Any(key.EndsWith);
                    if (exclude) continue;

                    var value = de.Value.ToString();

                    var r = rd.Data.NewRow();

                    r[Resx.KeyColumn] = key;

                    value = value.Replace("\r", "\\r");
                    value = value.Replace("\n", "\\n");

                    r[Resx.SourceTextColumn] = value;
                    rd.Data.Rows.Add(r);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("A problem occured reading " + neutralFile + "\n" + ex.Message, "Information",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            reader.Close();
        }

        private void AppendCulture(FileSystemInfo cultureFile, string culture, Resx rd, string[] excludeList)
        {
            var reader = new ResXResourceReader(cultureFile.FullName);

            try
            {
                foreach (DictionaryEntry de in reader)
                {
                    if (!(de.Value is string)) continue;

                    var key = (string)de.Key;

                    var exclude = excludeList.Any(key.EndsWith);
                    if (exclude) continue;

                    var value = de.Value.ToString();

                    var strWhere = String.Format("Key='{0}'", key);
                    var rows = rd.Data.Select(strWhere);
                    if (rows.Length == 0) throw new Exception("Row not Found");
                    var row = rows[0];
                    if (!rd.Data.Columns.Contains(culture))
                        rd.Data.Columns.Add(culture);

                    // update row
                    row.BeginEdit();

                    value = value.Replace("\r", "\\r");
                    value = value.Replace("\n", "\\n");
                    row[culture] = value;

                    row.EndEdit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("A problem occured reading " + cultureFile + "\n" + ex.Message, "Information",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            reader.Close();
        }

        private void FillCultures()
        {
            var array = CultureInfo.GetCultures(CultureTypes.AllCultures);
            Array.Sort(array, new CultureInfoComparer());
            foreach (var info in array)
            {
                if (!info.Equals(CultureInfo.InvariantCulture))
                {
                    listBoxCultures.Items.Add(info);
                }
            }

            var cList = Settings.Default.CultureList;

            foreach (var info in cList.Split(';').Where(c => !string.IsNullOrEmpty(c)).Select(cult => new CultureInfo(cult)))
            {
                listBoxSelected.Items.Add(info);
            }
        }

        private void AddCultures()
        {
            foreach (var t in listBoxCultures.SelectedItems)
            {
                var ci = (CultureInfo) t;

                if (listBoxSelected.Items.IndexOf(ci) == -1)
                    listBoxSelected.Items.Add(ci);
            }
        }

        private void AddExistingCultures()
        {
            if (string.IsNullOrEmpty(textBoxFolder.Text)) return;

            var files = Directory.GetFiles(textBoxFolder.Text, "*.resx");
            foreach (var f in files)
            {
                CultureInfo culture;
                if (!IsResxCultureSpecific(f, out culture)) continue;
                if (listBoxSelected.Items.IndexOf(culture) == -1)
                    listBoxSelected.Items.Add(culture);
            }
        }

        private void SaveCultures()
        {
            var cultures = String.Empty;
            foreach (var t in listBoxSelected.Items)
            {
                var info = (CultureInfo) t;

                if (cultures != String.Empty)
                    cultures = cultures + ";";

                cultures = cultures + info.Name;
            }

            Settings.Default.CultureList = cultures;
        }

        private void buttonBrowse_Click(object sender, EventArgs e)
        {
            if (resxDirectoryDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxFolder.Text = resxDirectoryDialog.SelectedPath;
                AddExistingCultures();
            }
        }

        private void buttonAdd_Click(object sender, EventArgs e)
        {
            AddCultures();
        }

        private void listBoxCultures_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            AddCultures();
        }

        private void buttonBrowseXls_Click(object sender, EventArgs e)
        {
            if (openFileDialogXls.ShowDialog() == DialogResult.OK)
            {
                textBoxXls.Text = openFileDialogXls.FileName;
            }
        }


        private void listBoxSelected_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listBoxSelected.SelectedItems.Count > 0)
            {
                listBoxSelected.Items.Remove(listBoxSelected.SelectedItems[0]);
            }
        }

        private void textBoxExclude_TextChanged(object sender, EventArgs e)
        {
            Settings.Default.ExcludeList = textBoxExclude.Text;
        }

        private void Resx2XlsForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveCultures();

            Settings.Default.Save();
        }

        private void textBoxFolder_TextChanged(object sender, EventArgs e)
        {
            Settings.Default.FolderPath = textBoxFolder.Text;
        }

        private void UpdateXls(string xlsFile, string projectRoot, string[] excludeList)
        {
            if (!File.Exists(xlsFile))
                return;

            var files = Directory.GetFiles(projectRoot, "*.resx");

            var rd = XlsToDataSet(xlsFile, null);

            foreach (var f in files)
            {
                var fi = new FileInfo(f);

                var fileRelativePath = fi.FullName.Remove(0, AddBS(projectRoot).Length);

                var reader = new ResXResourceReader(f) {BasePath = fi.DirectoryName};

                foreach (DictionaryEntry d in reader)
                {
                    if (!(d.Value is string)) continue;
                    var exclude = excludeList.Any(e => d.Key.ToString().EndsWith(e));
                    if (exclude) continue;

                    var strWhere = String.Format("FileSource ='{0}' AND Key='{1}'", fileRelativePath, d.Key);
                    var rows = rd.Data.Select(strWhere);

                    DataRow row;
                    if (rows.Length == 0)
                    {
                        // add row
                        row = rd.Data.NewRow();

                        // I update the neutral value
                        row["Key"] = d.Key.ToString();

                        rd.Data.Rows.Add(row);
                    }
                    else
                        row = rows[0];

                    // update row
                    row.BeginEdit();

                    var value = d.Value.ToString();
                    value = value.Replace("\r", "\\r");
                    value = value.Replace("\n", "\\n");
                    row["Value"] = value;

                    row.EndEdit();
                }
            }

            //delete unchenged rows
            foreach (DataRow r in rd.Data.Rows)
            {
                if (r.RowState == DataRowState.Unchanged)
                {
                    r.Delete();
                }
            }
            rd.Data.AcceptChanges();

            DataSetToXls(rd, new FileInfo(""), null);
        }

        private Resx XlsToDataSet(string xlsFile, string culture)
        {
            var app = new Application();
            var wb = app.Workbooks.Open(xlsFile,
                                             0, false, 5, "", "", false, XlPlatform.xlWindows, "",
                                             true, false, 0, true, false, false);

            var sheets = wb.Worksheets;

            var sheet = (Worksheet) sheets.Item[1];

            var rd = new Resx();

            var row = 3;

            while (true)
            {
                var fileSrc = ((Range) sheet.Cells[row, 1]).Text.ToString();

                if (String.IsNullOrEmpty(fileSrc))
                    break;

                var r = rd.Data.NewRow();

                r[Resx.KeyColumn] = ((Range)sheet.Cells[row, 1]).Text.ToString();
                r[Resx.SourceTextColumn] = ((Range)sheet.Cells[row, 2]).Text.ToString();
                r[Resx.UsageColumn] = ((Range) sheet.Cells[row, 3]).Text.ToString();

                var cult = ((Range) sheet.Cells[2, 4]).Text.ToString();

                if (String.IsNullOrEmpty(cult))
                    throw new Exception("Culture not found");

                if (!rd.Data.Columns.Contains(culture))
                    rd.Data.Columns.Add(culture);

                r[culture] = ((Range)sheet.Cells[row, 4]).Text.ToString();

                rd.Data.Rows.Add(r);

                row++;
            }

            rd.Data.AcceptChanges();

            wb.Close(false, objOpt, objOpt);
            app.Quit();

            return rd;
        }

        private void DataSetToXls(Resx rd, FileInfo destFile, string culture)
        {
            var app = new Application();
            var wb = app.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

            var sheets = wb.Worksheets;
            var sheet = (Worksheet) sheets.Item[1];
            sheet.Name = "Localize";

            sheet.Cells[1, 1] = Resx.KeyColumn;
            sheet.Cells[1, 2] = Resx.SourceTextColumn;
            sheet.Cells[1, 3] = Resx.UsageColumn;
            sheet.Cells[1, 4] = culture;

            var dw = rd.Data.DefaultView;
            dw.Sort = "Key";

            var row = 3;

            foreach (var r in from DataRowView drw in dw select drw.Row)
            {
                sheet.Cells[row, 1] = r[Resx.KeyColumn];
                sheet.Cells[row, 2] = r[Resx.SourceTextColumn];
                sheet.Cells[row, 3] = r[Resx.UsageColumn];
                sheet.Cells[row, 4] = r[culture];

                row++;
            }

            sheet.Cells.Range["A1", "Z1"].EntireColumn.AutoFit();

            // Save the Workbook and quit Excel.
            wb.SaveAs(destFile.FullName, objOpt, objOpt,
                      objOpt, objOpt, objOpt, XlSaveAsAccessMode.xlNoChange,
                      objOpt, objOpt, objOpt, objOpt, objOpt);
            wb.Close(false, objOpt, objOpt);
            app.Quit();
        }

        private IEnumerable<string> GetCulturesFromDataSet(Resx rd)
        {
            if (rd.Data.Rows.Count <= 0) return null;

            var columns = new DataColumn[rd.Data.Columns.Count];
            rd.Data.Columns.CopyTo(columns, 3);
            var cultureList = columns.Select(c => c.ColumnName).ToArray();

            return cultureList;
        }

        public static string JustStem(string cPath)
        {
            //Get the name of the file
            var lcFileName = JustFName(cPath.Trim());

            //Remove the extension and return the string
            if (lcFileName.IndexOf(".") == -1)
                return lcFileName;
            return lcFileName.Substring(0, lcFileName.LastIndexOf('.'));
        }

        public static string JustFName(string cFileName)
        {
            //Create the FileInfo object
            var fi = new FileInfo(cFileName);

            //Return the file name
            return fi.Name;
        }

        public static string AddBS(string cPath)
        {
            if (cPath.Trim().EndsWith("\\"))
            {
                return cPath.Trim();
            }
            return cPath.Trim() + "\\";
        }

        public void ShowXls(string xslFilePath)
        {
            if (!File.Exists(xslFilePath))
                return;

            var app = new Application();
            app.Workbooks.Open(xslFilePath,
                                             0, false, 5, "", "", false, XlPlatform.xlWindows, "",
                                             true, false, 0, true, false, false);

            app.Visible = true;
        }

        private void FinishWizard()
        {
            Cursor = Cursors.WaitCursor;

            var excludeList = textBoxExclude.Text.Split(';');

            var cultures = new string[listBoxSelected.Items.Count];
            for (var i = 0; i < listBoxSelected.Items.Count; i++)
            {
                cultures[i] = ((CultureInfo) listBoxSelected.Items[i]).Name;
            }

            switch (operation)
            {
                case ResxToXlsOperation.Export:

                    if (String.IsNullOrEmpty(textBoxFolder.Text))
                    {
                        MessageBox.Show(
                            "You must select a the .Net Project root wich contains your updated resx files", "Update",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);

                        wizardControl1.CurrentStepIndex = intermediateStepProject.StepIndex;

                        return;
                    }

                    if (xlsDirectoryDialog.ShowDialog() == DialogResult.OK)
                    {
                        System.Windows.Forms.Application.DoEvents();

                        var path = xlsDirectoryDialog.SelectedPath;

                        ResxToXls(textBoxFolder.Text, path, cultures, excludeList);
                    }
                    break;
                case ResxToXlsOperation.ImportFile:
                    if (String.IsNullOrEmpty(textBoxXls.Text))
                    {
                        MessageBox.Show("You must select the Excel document to update", "Update", MessageBoxButtons.OK,
                                        MessageBoxIcon.Information);

                        wizardControl1.CurrentStepIndex = intermediateStepXlsSelect.StepIndex;

                        return;
                    }

                    XlsToResx(textBoxXls.Text);

                    MessageBox.Show("Localized Resources created.", "Build", MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);

                    break;
                case ResxToXlsOperation.ImportDirectory:
                    if (String.IsNullOrEmpty(textBoxFolder.Text))
                    {
                        MessageBox.Show(
                            "You must select a the .Net Project root wich contains your updated resx files", "Update",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);

                        wizardControl1.CurrentStepIndex = intermediateStepProject.StepIndex;

                        return;
                    }

                    if (String.IsNullOrEmpty(textBoxXls.Text))
                    {
                        MessageBox.Show("You must select the Excel document to update", "Update", MessageBoxButtons.OK,
                                        MessageBoxIcon.Information);

                        wizardControl1.CurrentStepIndex = intermediateStepXlsSelect.StepIndex;

                        return;
                    }


                    UpdateXls(textBoxXls.Text, textBoxFolder.Text, excludeList);

                    MessageBox.Show("Excel Document Updated.", "Update", MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                    break;
            }

            Cursor = Cursors.Default;

            Close();
        }

        private void wizardControl1_CurrentStepIndexChanged(object sender, EventArgs e)
        {
        }

        private void wizardControl1_NextButtonClick(WizardControl sender, WizardNextButtonClickEventArgs args)
        {
            var index = wizardControl1.CurrentStepIndex;

            var offset = 1;

            switch (index)
            {
                case 0:

                    switch (operation)
                    {
                        case ResxToXlsOperation.Export:
                            wizardControl1.CurrentStepIndex = 1 - offset;
                            break;
                        case ResxToXlsOperation.ImportFile:
                            wizardControl1.CurrentStepIndex = 4 - offset;
                            break;
                        case ResxToXlsOperation.ImportDirectory:
                            wizardControl1.CurrentStepIndex = 1 - offset;
                            break;
                    }
                    break;

                case 1:

                    switch (operation)
                    {
                        case ResxToXlsOperation.ImportDirectory:
                            wizardControl1.CurrentStepIndex = 4 - offset;
                            break;
                    }
                    break;


                case 3:

                    switch (operation)
                    {
                        case ResxToXlsOperation.Export:
                            wizardControl1.CurrentStepIndex = 5 - offset;
                            break;
                    }
                    break;
            }
        }

        private void wizardControl1_BackButtonClick(WizardControl sender, WizardClickEventArgs args)
        {
            var index = wizardControl1.CurrentStepIndex;

            var offset = 1;

            switch (index)
            {
                case 5:

                    switch (operation)
                    {
                        case ResxToXlsOperation.Export:
                            wizardControl1.CurrentStepIndex = 3 + offset;
                            break;
                    }
                    break;
                case 4:

                    switch (operation)
                    {
                        case ResxToXlsOperation.ImportFile:
                            wizardControl1.CurrentStepIndex = 0 + offset;
                            break;
                        case ResxToXlsOperation.ImportDirectory:
                            wizardControl1.CurrentStepIndex = 1 + offset;
                            break;
                    }
                    break;
            }
        }

        private void wizardControl1_FinishButtonClick(object sender, EventArgs e)
        {
            FinishWizard();
        }

        private void startStep1_Click(object sender, EventArgs e)
        {
        }

        private enum ResxToXlsOperation
        {
            Export,
            ImportFile,
            ImportDirectory
        };

        private void radioButtonCreateXls_CheckedChanged(object sender, EventArgs e)
        {

        }
    }
}