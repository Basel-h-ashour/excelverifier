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

        private void MainForm_Load(object sender, EventArgs e)
        {
            inputFormat.SelectedIndex = outputFormat.SelectedIndex = 0;
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
                excelData = ExcelOperations.GetExcelFileAsOneDataSetTable(c);

                if (excelData == null)
                {
                    MessageBox.Show(inputTextBox.Text + "\n\nFailed to read excel file.", "Empty or Corrupted File", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                outputExcelData = new DataSet(excelData.DataSetName + " VALIDATED");

                Dictionary<string, string> columnMapper = new Dictionary<string, string>();
                HashSet<string> requiredFields = new HashSet<string>();

                if (outputFormat.SelectedItem.ToString() == "VIA MEDICA" && inputFormat.SelectedItem.ToString() == "VIA MEDICA")
                {
                    columnMapper.Add("EMP No / UAEID", "EMP No / UAEID");
                    columnMapper.Add("Name", "Name");
                    columnMapper.Add("Date_Of_Birth", "Date_Of_Birth");
                    columnMapper.Add("Emirates_ID", "Emirates_ID");
                    columnMapper.Add("NATIONALITY", "NATIONALITY");
                    columnMapper.Add("Gender", "Gender");
                    columnMapper.Add("Contact", "Contact");
                    columnMapper.Add("Passport_Number", "Passport_Number");
                    columnMapper.Add("address", "address");
                    columnMapper.Add("Emirate", "Emirate");
                    columnMapper.Add("Referring_Facility_MRN", "Referring_Facility_MRN");
                    columnMapper.Add("Referring_Facility_ID", "Referring_Facility_ID");
                    columnMapper.Add("Referring_Facility_Name", "Referring_Facility_Name");
                    columnMapper.Add("Ordering_Physician", "Ordering_Physician");
                    columnMapper.Add("Hasana ID", "Hasana ID");

                    requiredFields.Add("Gender");
                    requiredFields.Add("Name");
                    requiredFields.Add("Contact");
                    requiredFields.Add("NATIONALITY");
                    requiredFields.Add("Date_Of_Birth");
                    requiredFields.Add("Referring_Facility_MRN");
                }
                else if (outputFormat.SelectedItem.ToString() == "LDM" && inputFormat.SelectedItem.ToString() == "VIA MEDICA")
                {
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

                    requiredFields.Add("GENDER");
                    requiredFields.Add("PATIENT NAME");
                    requiredFields.Add("Mobile Number");
                    requiredFields.Add("NATIONALITY");
                    requiredFields.Add("BIRTHDATE");
                    requiredFields.Add("MRN");
                }


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

                    //A Dictionary storing the data that should be autocompleted for each applicable column
                    Dictionary<string, string> defaultAutocompletion = new Dictionary<string, string>();

                    //A set for each field that can't be duplicated (stores the duplicates)
                    Dictionary<string, HashSet<string>> duplicates = new Dictionary<string, HashSet<string>>();

                    //A set that detects the duplicates when keys repeat (only detects duplicates, and sends to above sets)
                    Dictionary<string, HashSet<string>> duplicatesDetector = new Dictionary<string, HashSet<string>>();

                    //A set for each field that can't be combine duplicated (stores the duplicates)
                    Dictionary<string, Dictionary<string, int>> duplicatesCombined = new Dictionary<string, Dictionary<string, int>>();

                    //A set that detects the combined duplicates when keys repeat (only detects duplicates, and sends to above sets)
                    Dictionary<string, HashSet<string>> duplicatesDetectorCombined = new Dictionary<string, HashSet<string>>();

                    //a set of fields common between all formats and having specific checks/modification (unifying their names with a known Key)
                    Dictionary<string, string> fieldNames = new Dictionary<string, string>();

                    outputExcelData.Tables.Add(excelData.Tables[i].TableName);

                    //Setting up output format for sheet + default autocompletion keys + duplicate detectors
                    switch (outputFormat.SelectedItem.ToString())
                    {
                        case "VIA MEDICA":
                            outputExcelData.Tables[i].Columns.Add("EMP No / UAEID");
                            outputExcelData.Tables[i].Columns.Add("Name");
                            outputExcelData.Tables[i].Columns.Add("Date_Of_Birth");
                            outputExcelData.Tables[i].Columns.Add("Emirates_ID");
                            outputExcelData.Tables[i].Columns.Add("NATIONALITY");
                            outputExcelData.Tables[i].Columns.Add("Gender");
                            outputExcelData.Tables[i].Columns.Add("Contact");
                            outputExcelData.Tables[i].Columns.Add("Passport_Number");
                            outputExcelData.Tables[i].Columns.Add("address");
                            outputExcelData.Tables[i].Columns.Add("Emirate");
                            outputExcelData.Tables[i].Columns.Add("Referring_Facility_MRN");
                            outputExcelData.Tables[i].Columns.Add("Referring_Facility_ID");
                            outputExcelData.Tables[i].Columns.Add("Referring_Facility_Name");
                            outputExcelData.Tables[i].Columns.Add("Ordering_Physician");
                            outputExcelData.Tables[i].Columns.Add("Hasana ID");
                            outputExcelData.Tables[i].Columns.Add("Notes");

                            fieldNames.Add("passport_number", "Passport_Number");
                            fieldNames.Add("EID", "Emirates_ID");
                            fieldNames.Add("gender", "Gender");
                            fieldNames.Add("MRN", "Referring_Facility_MRN");
                            fieldNames.Add("mobile_number", "Contact");

                            defaultAutocompletion.Add("Emirate", "");

                            duplicates.Add("Passport_Number", new HashSet<string>());
                            duplicates.Add("Emirates_ID", new HashSet<string>());
                            duplicates.Add("Referring_Facility_MRN", new HashSet<string>());

                            duplicatesCombined.Add("Emirates_ID|Referring_Facility_MRN", new Dictionary<string, int>());


                            duplicatesDetector.Add("Passport_Number", new HashSet<string>());
                            duplicatesDetector.Add("Emirates_ID", new HashSet<string>());
                            duplicatesDetector.Add("Referring_Facility_MRN", new HashSet<string>());

                            duplicatesDetectorCombined.Add("Emirates_ID|Referring_Facility_MRN", new HashSet<string>());

                            break;

                        case "LDM":
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

                            fieldNames.Add("passport_number", "PASSPORT NUMBER");
                            fieldNames.Add("EID", "EMIRATES ID");
                            fieldNames.Add("gender", "GENDER");
                            fieldNames.Add("MRN", "MRN");
                            fieldNames.Add("mobile_number", "Mobile Number");

                            defaultAutocompletion.Add("city", "");

                            duplicates.Add("PASSPORT NUMBER", new HashSet<string>());
                            duplicates.Add("EMIRATES ID", new HashSet<string>());
                            duplicates.Add("MRN", new HashSet<string>());

                            duplicatesCombined.Add("EMIRATES ID|MRN", new Dictionary<string, int>());

                            duplicatesDetector.Add("PASSPORT NUMBER", new HashSet<string>());
                            duplicatesDetector.Add("EMIRATES ID", new HashSet<string>());

                            duplicatesDetectorCombined.Add("EMIRATES ID|MRN", new HashSet<string>());

                            break;
                    }

                    for (int j = 0; j < excelData.Tables[i].Rows.Count; j++)
                    {
                        if (excelData.Tables[i].Rows[j].ItemArray.All(v => v.ToString() == ""))
                        {
                            excelData.Tables[i].Rows[j].Delete();
                            continue;
                        }

                        foreach (var detector in duplicatesDetector)
                        {
                            string cellData = excelData.Tables[i].Rows[j][columnMapper[detector.Key]].ToString();

                            if (duplicatesDetector[detector.Key].Contains(cellData))
                            {
                                duplicates[detector.Key].Add(cellData);
                            }
                            else
                            {
                                duplicatesDetector[detector.Key].Add(cellData);
                            }
                        }

                        foreach (var detector in duplicatesDetectorCombined)
                        {
                            string[] columns = detector.Key.Split('|');
                            string cellData = "";

                            for (int k = 0; k < columns.Length; k++)
                            {
                                cellData += excelData.Tables[i].Rows[j][columns[k]].ToString();
                            }

                            if (duplicatesDetectorCombined[detector.Key].Contains(cellData))
                            {
                                if (duplicatesCombined[detector.Key].ContainsKey(cellData))
                                    duplicatesCombined[detector.Key][cellData]++;
                                else
                                    duplicatesCombined[detector.Key].Add(cellData, 2);
                            }
                            else
                            {
                                duplicatesDetectorCombined[detector.Key].Add(cellData);
                            }
                        }

                        for (int k = 0; k < defaultAutocompletion.Count; k++)
                        {
                            string cellData = excelData.Tables[i].Rows[j][columnMapper[defaultAutocompletion.ElementAt(k).Key]].ToString();

                            if (defaultAutocompletion.ContainsKey(defaultAutocompletion.ElementAt(k).Key) && defaultAutocompletion[defaultAutocompletion.ElementAt(k).Key] == "" && cellData != "")
                            {
                                defaultAutocompletion[defaultAutocompletion.ElementAt(k).Key] = cellData;
                            }
                        }
                    }

                    excelData.AcceptChanges();

                    for (int j = 0; j < excelData.Tables[i].Rows.Count; j++)
                    {
                        //Before adding the output row, check first if both EID and Ref Number are duplicated
                        if (duplicatesCombined[fieldNames["EID"] + "|" + fieldNames["MRN"]].ContainsKey(excelData.Tables[i].Rows[j][fieldNames["EID"]].ToString() + excelData.Tables[i].Rows[j][fieldNames["MRN"]].ToString()))
                        {
                            if (duplicatesCombined[fieldNames["EID"] + "|" + fieldNames["MRN"]][excelData.Tables[i].Rows[j][fieldNames["EID"]].ToString() + excelData.Tables[i].Rows[j][fieldNames["MRN"]].ToString()] > 1)
                            {
                                duplicatesCombined[fieldNames["EID"] + "|" + fieldNames["MRN"]][excelData.Tables[i].Rows[j][fieldNames["EID"]].ToString() + excelData.Tables[i].Rows[j][fieldNames["MRN"]].ToString()]--;
                                excelData.Tables[i].Rows[j].Delete();
                                excelData.AcceptChanges();
                                j--;
                                continue;
                            }
                        }

                        outputExcelData.Tables[i].Rows.Add();

                        //check if that at least NID or Passport No. is available (use input column name here ONLY)
                        if (String.IsNullOrEmpty(excelData.Tables[i].Rows[j][columnMapper[fieldNames["EID"]]].ToString()) && String.IsNullOrEmpty(excelData.Tables[i].Rows[j][columnMapper[fieldNames["passport_number"]]].ToString()))
                        {
                            outputExcelData.Tables[i].Rows[j]["Notes"] = outputExcelData.Tables[i].Rows[j]["Notes"].ToString() + "At least EID or Passport No. is required and neither is entered; ";
                        }

                        for (int k = 0; k < outputExcelData.Tables[i].Columns.Count - 1; k++)
                        {
                            string outputColumnName = outputExcelData.Tables[i].Columns[k].ColumnName;

                            string cellData;

                            //skip with a blank if column only exists in output excel format and is not SN (which is autogenerated)
                            if (!columnMapper.ContainsKey(outputColumnName))
                            {
                                cellData = outputColumnName == "SN" ? (j + 1).ToString() : "";
                                outputExcelData.Tables[i].Rows[j][outputColumnName] = cellData;
                                continue;
                            }

                            cellData = excelData.Tables[i].Rows[j][columnMapper[outputColumnName]].ToString();

                            //check if column is required but left blank
                            if (String.IsNullOrEmpty(cellData) && requiredFields.Contains(outputColumnName))
                            {
                                outputExcelData.Tables[i].Rows[j]["Notes"] = outputExcelData.Tables[i].Rows[j]["Notes"].ToString() + columnMapper[outputColumnName] + " is required but not entered; ";
                                outputExcelData.Tables[i].Rows[j][outputColumnName] = cellData;
                                continue;

                            }

                            //Check for duplicates in the fields marked before to be unique
                            if (cellData != "" && duplicates.ContainsKey(outputColumnName) && duplicates[outputColumnName].Contains(cellData))
                            {
                                outputExcelData.Tables[i].Rows[j]["Notes"] = outputExcelData.Tables[i].Rows[j]["Notes"].ToString() + columnMapper[outputColumnName] + " is duplicated; ";
                            }

                            //If column is marked to be autocompleted, use the default autocompletion found in first run
                            cellData = defaultAutocompletion.ContainsKey(outputColumnName) && String.IsNullOrEmpty(cellData) ? defaultAutocompletion[outputColumnName] : cellData;

                            if (outputFormat.SelectedItem.ToString() == "LDM")
                            {
                                //Only use initials for gender
                                cellData = outputColumnName == fieldNames["gender"] ? cellData.Substring(0, 1) : cellData;

                                //Remove dashes from EID
                                cellData = outputColumnName == fieldNames["EID"] ? cellData.Replace("-", "") : cellData;

                                //Remove country code from mobile number
                                cellData = outputColumnName == fieldNames["mobile_number"] ? cellData.Substring(6).Replace("-", "") : cellData;
                            }

                            outputExcelData.Tables[i].Rows[j][outputColumnName] = cellData;
                        }
                    }

                    outputExcelData.AcceptChanges();
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
