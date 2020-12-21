using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ExcelVerifier
{
    public partial class MainForm : Form
    {
        public DataSet outputExcelData;

        public MainForm()
        {
            InitializeComponent();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            Application.Exit();
        }

        private void validateButton_Click(object sender, EventArgs e)
        {
            GC.Collect();

            if (!File.Exists(inputTextBox.Text))
            {
                MessageBox.Show("Input File does not exist.", "File Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (ExcelOperations.FileIsOpened(new FileInfo(inputTextBox.Text)))
            {
                MessageBox.Show(inputTextBox.Text + "\n\nInput Excel file is currently in use.\nPlease close it and try again.", "File In Use", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            statusLabel.Text = "STATUS:  WORKING. PLEASE WAIT...";
            statusLabel.ForeColor = System.Drawing.Color.Yellow;
            statusLabel.Refresh();

            DataSet excelData;

            using (var c = ExcelOperations.GetConnection(inputTextBox.Text, true))
            {
                excelData = ExcelOperations.GetExcelFileAsDataSet(c);

                outputExcelData = new DataSet(excelData.DataSetName + " VALIDATED");

                Dictionary<string, string> columnMapper = new Dictionary<string, string>();
                columnMapper.Add("PATIENT NAME", "Name");
                columnMapper.Add("BIRTHDATE", "Date_Of_Birth");
                columnMapper.Add("EMIRATES ID", "Emirates_ID");
                columnMapper.Add("NATIONALITY", "NATIONALITY");
                columnMapper.Add("GENDER", "Gender");
                columnMapper.Add("Mobile Number", "Contact");
                columnMapper.Add("Address", "address");
                columnMapper.Add("MRN", "Referring_Facility_MRN");
                columnMapper.Add("PASSPORT NUMBER", "Passport_Number");
                columnMapper.Add("city", "Emirate");

                HashSet<string> requiredFields = new HashSet<string>();
                requiredFields.Add("GENDER");
                requiredFields.Add("PATIENT NAME");
                requiredFields.Add("Mobile Number");
                requiredFields.Add("NATIONALITY");
                requiredFields.Add("BIRTHDATE");

                //Finding the duplicate values in each of the unique fields AND finding the default autocompletions
                for (int i = 0; i < excelData.Tables.Count; i++)
                {
                    //check if all mapped columns exist in source excel file
                    foreach (KeyValuePair<string, string> column in columnMapper)
                    {
                        if (!excelData.Tables[i].Columns.Contains(column.Value.ToLower()))
                        {
                            statusLabel.Text = "STATUS:  ERROR!";
                            statusLabel.ForeColor = System.Drawing.Color.Red;
                            MessageBox.Show("\"" + column.Value + "\" column should be present in the source excel sheet \"" + excelData.Tables[i].TableName + "\".\nPlease check the format and try again.", "Missing Column", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            outputExcelData = null;
                            return;
                        }
                    }

                    outputExcelData.Tables.Add();

                    //Setting up output columns for each sheet
                    outputExcelData.Tables[i].Columns.Add("SN");
                    outputExcelData.Tables[i].Columns.Add("SAMPLE ID");
                    outputExcelData.Tables[i].Columns.Add("RACK NO.");
                    outputExcelData.Tables[i].Columns.Add("EMIRATES ID");
                    outputExcelData.Tables[i].Columns.Add("MRN");
                    outputExcelData.Tables[i].Columns.Add("PATIENT NAME");
                    outputExcelData.Tables[i].Columns.Add("BIRTHDATE");
                    outputExcelData.Tables[i].Columns.Add("GENDER");
                    outputExcelData.Tables[i].Columns.Add("COLLECTION DATE AND  TIME");
                    outputExcelData.Tables[i].Columns.Add("PASSPORT NUMBER");
                    outputExcelData.Tables[i].Columns.Add("NATIONALITY");
                    outputExcelData.Tables[i].Columns.Add("HASANA ID");
                    outputExcelData.Tables[i].Columns.Add("AI (Additional Identifier)");
                    outputExcelData.Tables[i].Columns.Add("Mobile Number");
                    outputExcelData.Tables[i].Columns.Add("email");
                    outputExcelData.Tables[i].Columns.Add("city");
                    outputExcelData.Tables[i].Columns.Add("Address");
                    outputExcelData.Tables[i].Columns.Add("Notes");

                    //A Dictionary storing the data that should be autocompleted for each applicable column
                    Dictionary<string, string> defaultAutocompletion = new Dictionary<string, string>();
                    defaultAutocompletion.Add("Emirate", "");

                    //A set for each field that can't be duplicated (stores the duplicates)
                    Dictionary<string, HashSet<string>> duplicates = new Dictionary<string, HashSet<string>>();
                    duplicates.Add("Passport_Number", new HashSet<string>());
                    duplicates.Add("Emirates_ID", new HashSet<string>());
                    duplicates.Add("Referring_Facility_MRN", new HashSet<string>());


                    //A set that detects the duplicates when keys repeat (only detects duplicates, and sends to above sets)
                    Dictionary<string, HashSet<string>> duplicatesDetector = new Dictionary<string, HashSet<string>>();
                    duplicatesDetector.Add("Passport_Number", new HashSet<string>());
                    duplicatesDetector.Add("Emirates_ID", new HashSet<string>());
                    duplicatesDetector.Add("Referring_Facility_MRN", new HashSet<string>());

                    for (int j = 0; j < excelData.Tables[i].Rows.Count; j++)
                    {
                        bool allBlank = true;

                        for (int k = 0; k < excelData.Tables[i].Columns.Count; k++)
                        {
                            string columnName = excelData.Tables[i].Columns[k].ColumnName;
                            string cellData = excelData.Tables[i].Rows[j][k].ToString();

                            if (!String.IsNullOrEmpty(cellData)) allBlank = false;

                            if (defaultAutocompletion.ContainsKey(columnName) && defaultAutocompletion[columnName] != "" && cellData != "")
                            {
                                defaultAutocompletion[columnName] = cellData;
                            }

                            if (duplicatesDetector.ContainsKey(columnName))
                            {
                                if (duplicatesDetector[columnName].Contains(cellData))
                                {
                                    duplicates[columnName].Add(cellData);
                                }
                                else
                                {
                                    duplicatesDetector[columnName].Add(cellData);
                                }
                            }
                        }

                        if (allBlank) excelData.Tables[i].Rows[j].Delete();
                    }

                    excelData.AcceptChanges();

                    for (int j = 0; j < excelData.Tables[i].Rows.Count; j++)
                    {
                        if (excelData.Tables[i].Rows[j].RowState == DataRowState.Modified) continue;

                        outputExcelData.Tables[i].Rows.Add();

                        //check if that at least NID or Passport No. is available
                        if (String.IsNullOrEmpty(excelData.Tables[i].Rows[j]["Emirates_ID"].ToString()) && String.IsNullOrEmpty(excelData.Tables[i].Rows[j]["Passport_Number"].ToString()))
                        {
                            outputExcelData.Tables[i].Rows[j]["Notes"] = outputExcelData.Tables[i].Rows[j]["Notes"].ToString() + "At least EID or Passport No. is required and neither is entered; ";
                        }

                        for (int k = 0; k < outputExcelData.Tables[i].Columns.Count - 1; k++)
                        {
                            string outputColumnName = outputExcelData.Tables[i].Columns[k].ColumnName;

                            string cellData;
                            string inputcolumnName;

                            //skip with a blank if column only exists in output excel format
                            if (!columnMapper.ContainsKey(outputColumnName))
                            {
                                cellData = outputColumnName == "SN" ? (j + 1).ToString() : "";
                                outputExcelData.Tables[i].Rows[j][outputColumnName] = cellData;
                                continue;
                            }

                            inputcolumnName = columnMapper[outputColumnName];
                            cellData = excelData.Tables[i].Rows[j][inputcolumnName].ToString();

                            //check if column is required but left blank
                            if (String.IsNullOrEmpty(cellData) && requiredFields.Contains(outputColumnName))
                            {
                                outputExcelData.Tables[i].Rows[j]["Notes"] = outputExcelData.Tables[i].Rows[j]["Notes"].ToString() + inputcolumnName + " is required but not entered; ";
                                outputExcelData.Tables[i].Rows[j][k] = cellData;
                                continue;

                            }

                            //Check for duplicates in the fields marked before to be unique
                            if (cellData != "" && duplicates.ContainsKey(inputcolumnName) && duplicates[inputcolumnName].Contains(cellData))
                            {
                                outputExcelData.Tables[i].Rows[j]["Notes"] = outputExcelData.Tables[i].Rows[j]["Notes"].ToString() + inputcolumnName + " is duplicated; ";
                            }

                            //If column is marked to be autocompleted, use the default autocompletion found in first run
                            cellData = defaultAutocompletion.ContainsKey(inputcolumnName) && String.IsNullOrEmpty(cellData) ? defaultAutocompletion[inputcolumnName] : cellData;

                            //Only use initials for gender
                            cellData = outputColumnName == "GENDER" ? cellData.Substring(0, 1) : cellData;

                            //Remove dashes from EID
                            cellData = outputColumnName == "EMIRATES ID" ? cellData.Replace("-", "") : cellData;

                            //Remove country code from mobile number
                            cellData = outputColumnName == "Mobile Number" ? cellData.Substring(6).Replace("-", "") : cellData;


                            outputExcelData.Tables[i].Rows[j][k] = cellData;
                        }
                    }
                }
            }

            statusLabel.Text = "STATUS:   SUCESSFULLY LOADED " + inputTextBox.Text.Split('\\').Last();
            statusLabel.ForeColor = System.Drawing.Color.GreenYellow;
            statusLabel.Refresh();
        }

        private void inputBrowseButton_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (dialog.ShowDialog() != DialogResult.OK) return;

            string fileExtension = dialog.FileName.Split('.').Last();

            if (fileExtension != "csv" && fileExtension != "xlsx" && fileExtension != "xls")
            {
                MessageBox.Show("File must be in an excel format.", "Invalid File Type", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            inputTextBox.Text = dialog.FileName;
            outputTextBox.Text = dialog.FileName.Replace(dialog.FileName.Split('\\').Last(), "");
        }

        private void outputBrowseButton_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();

            if (dialog.ShowDialog() != DialogResult.OK) return;


            outputTextBox.Text = dialog.SelectedPath;
        }

        private void exportButton_Click(object sender, EventArgs e)
        {
            if (outputExcelData == null)
            {
                MessageBox.Show("\nPlease validate an excel file first before exporting", "No Excel Validated", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            string outputFileName = inputTextBox.Text.Split('\\').Last().Replace(".xlsx", " - VALIDATED.xlsx");
            string outputFilePath = outputTextBox.Text.EndsWith("\\") ? outputTextBox.Text + outputFileName : outputTextBox.Text + "\\" + outputFileName;

            if (String.IsNullOrEmpty(outputTextBox.Text) || !Directory.Exists(outputTextBox.Text))
            {
                MessageBox.Show("Output Directory does not exist.", "Directory Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (File.Exists(outputFilePath) && ExcelOperations.FileIsOpened(new FileInfo(outputFilePath)))
            {
                    MessageBox.Show(outputFileName + "\nOutput excel file is currently in use. Please close it and try again.", "File In Use", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
            }

            statusLabel.Text = "STATUS:  EXPORTING " + inputTextBox.Text.Split('\\').Last();
            statusLabel.ForeColor = System.Drawing.Color.Yellow;
            statusLabel.Refresh();

            CreateExcelFile.CreateExcelDocument(outputExcelData, outputFilePath);
            
            statusLabel.Text = "STATUS:  SUCCESSFULLY EXPORTED " + inputTextBox.Text.Split('\\').Last();
            statusLabel.ForeColor = System.Drawing.Color.Green;
            statusLabel.Refresh();
        }
    }
}
