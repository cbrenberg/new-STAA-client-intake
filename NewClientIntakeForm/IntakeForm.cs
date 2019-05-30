using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;

using NewClientIntakeForm.Models;

namespace NewClientIntakeForm
{
    public partial class NewClientIntakeForm : Form
    {
        private NamedEntity SigningAttorney;
        private Complainant Complainant;
        private NamedEntityWithAddress RespondentCompanyOne;
        private NamedEntityWithAddress RespondentCompanyTwo;
        private NamedEntity RespondentIndividualOne;
        private NamedEntity RespondentIndividualTwo;
        private OSHARegion OSHA;

        private string FullPathToClientDirectory;

        private Dictionary<string, string> BookmarkList = new Dictionary<string, string>();

        
        public NewClientIntakeForm()
        {
            InitializeComponent();
        }

        private void BtnSubmit_Click(object sender, EventArgs e)
        {
            this.Enabled = false;

            CreateEntitiesFromFormData();

            CreateClientDirectoryStructure();

            WriteCaseInfoTextFile();

            PopulateBookmarkList();

            CreateTemplatesForSelectedItems();

            this.Close();
            Application.Exit();
        }

        private void CreateEntitiesFromFormData()
        {
            Complainant = new Complainant(
                (IsMissing(CompFirstName.Text) && IsMissing(CompLastName.Text)) ? "COMPLAINANT NAME" : (CompFirstName.Text.Trim() + " " + CompLastName.Text.Trim()),
                new Address(
                    IsMissing(CompAddLine1.Text) ? "COMPLAINANT ADDRESS LINE 1" : CompAddLine1.Text,
                    IsMissing(CompAddLine2.Text) ? "COMPLAINANT ADDRESS LINE 2" : CompAddLine2.Text
                    ),
                DateOfHirePicker.Value,
                DateOfTerminationPicker.Value,
                IsMissing(CompEmail.Text) ? "COMPLAINANT EMAIL" : CompEmail.Text,
                IsMissing(CompPhone.Text) ? "COMPLAINANT PHONE" : CompPhone.Text
                );
            Complainant.HasDefaultValues = IsMissing(CompFirstName.Text) ? true : false;

            RespondentCompanyOne = new NamedEntityWithAddress(
                IsMissing(RespComp1Name.Text) ? "RESPONDENT COMPANY 1" : RespComp1Name.Text,
                new Address(
                    IsMissing(RespComp1AddLine1.Text) ? "RESPONDENT 1 ADDRESS LINE 1" : RespComp1AddLine1.Text,
                    IsMissing(RespComp1AddLine2.Text) ? "RESPONDENT 1 ADDRESS LINE 2" : RespComp1AddLine2.Text
                    )
                );
            RespondentCompanyOne.HasDefaultValues = IsMissing(RespComp1Name.Text) ? true : false;

            RespondentCompanyTwo = new NamedEntityWithAddress(
                IsMissing(RespComp2Name.Text) ? "RESPONDENT COMPANY 2" : RespComp2Name.Text,
                new Address(
                    IsMissing(RespComp2AddLine1.Text) ? "RESPONDENT 2 ADDRESS LINE 1" : RespComp2AddLine1.Text,
                    IsMissing(RespComp2AddLine2.Text) ? "RESPONDENT 2 ADDRESS LINE 2" : RespComp2AddLine2.Text
                    )
                );
            RespondentCompanyTwo.HasDefaultValues = IsMissing(RespComp2Name.Text) ? true : false;

            RespondentIndividualOne = new NamedEntity(IsMissing(RespInd1Name.Text) ? "RESPONDENT INDIVIDUAL 1" : RespInd1Name.Text);
            RespondentIndividualOne.HasDefaultValues = IsMissing(RespInd1Name.Text) ? true : false;

            RespondentIndividualTwo = new NamedEntity(IsMissing(RespInd2Name.Text) ? "RESPONDENT INDIVIDUAL 2" : RespInd2Name.Text);
            RespondentIndividualTwo.HasDefaultValues = IsMissing(RespInd2Name.Text) ? true : false;

            //NamedEntity[] AllRespondents = new NamedEntity[] { RespondentCompanyOne, RespondentCompanyTwo, RespondentIndividualOne, RespondentIndividualTwo };

            SigningAttorney = new NamedEntity(IsMissing(AttorneyName.Text) ? "SIGNING ATTORNEY" : AttorneyName.Text);

            OSHA = new OSHARegion(selectOSHA.Text);
        }

        private void CreateClientDirectoryStructure()
        {
            toolStripStatusLabel1.Text = "Creating client folder...";

            //create client folder structure
            try
            {
                string currentDirectory = Directory.GetCurrentDirectory();
                //string rootDriveLetter = Directory.GetDirectoryRoot(currentDirectory);
                //for debugging:
                string rootDriveLetter = "X:\\";
                string directoryToSearch = rootDriveLetter + "Shared\\Active Clients";
                string[] oshaFolder = Directory.GetDirectories(directoryToSearch, "OSHA");
                string basepath = oshaFolder[0];
                string clientFolderName = $"{CompLastName.Text.Trim()}, {CompFirstName.Text.Trim()}";
                FullPathToClientDirectory = basepath + "\\" + clientFolderName;
                if (!Directory.Exists(FullPathToClientDirectory))
                {
                    //create client folder in OSHA folder
                    Directory.CreateDirectory(FullPathToClientDirectory);
                }
                else
                {
                    MessageBox.Show("File already exists for this client. Please fill templates manually.");
                    Application.Exit();
                }

            }
            catch (Exception exc)
            {
                Console.Write("Could not create directory structure: ", exc);
            }
        }

        private void CreateTemplatesForSelectedItems()
        {
            var checkedItems = SelectedDocsCheckboxList.CheckedItems;
            KeyValuePair<string, string>[] listOfTemplatesToPopulate = new KeyValuePair<string, string>[checkedItems.Count];
            for (int i = 0; i < checkedItems.Count; i++)
            {
                switch (checkedItems[i].ToString())
                {
                    case "Anatomy of a Lawsuit":
                        listOfTemplatesToPopulate[i] = new KeyValuePair<string, string>("Anatomy Letter Template.dotx", "Anatomy of a Lawsuit Draft v1.docx");
                        break;
                    case "Retainer Agreement":
                        listOfTemplatesToPopulate[i] = new KeyValuePair<string, string>("Retainer Template.dotx", "Retainer Draft v1.docx");
                        break;
                    case "OSHA Complaint":
                        listOfTemplatesToPopulate[i] = new KeyValuePair<string, string>("New Complaint Template.dotx", "Complaint Draft v1.docx");
                        break;
                    case "Settlement Agreement":
                        break;
                    default:
                        break;
                }

            }

            foreach (KeyValuePair<string, string> template in listOfTemplatesToPopulate)
            {
                toolStripProgressBar1.Value += 100 / listOfTemplatesToPopulate.Length;
                PopulateClientTemplate(template);
            }
        }

        private void WriteCaseInfoTextFile()
        {
            string clientInfoPath = FullPathToClientDirectory + "\\Client Information.txt";
            string clientText = "Complainant\r\n--------------------\r\nName: " + Complainant.Name + "\r\nAddress:\r\n" + Complainant.Address.Line1 + "\r\n" + Complainant.Address.Line2 + "\r\nPhone: " + Complainant.PhoneNumber + "\r\nEmail: " + Complainant.EmailAddress + "\r\nDate Of Hire: " + Complainant.DateOfHire.ToString("D") + "\r\nDate Of Termination: " + Complainant.DateOfTermination.ToString("D") + "\r\n";
            string company1Text = "Respondent Company 1\r\n--------------------\r\nName: " + RespondentCompanyOne.Name + "\r\nAddress:\r\n" + RespondentCompanyOne.Address.Line1 + "\r\n" + RespondentCompanyOne.Address.Line2 + "\r\n";
            string company2Text = "Respondent Company 2\r\n--------------------\r\nName: " + RespondentCompanyTwo.Name + "\r\nAddress:\r\n" + RespondentCompanyTwo.Address.Line1 + "\r\n" + RespondentCompanyTwo.Address.Line2 + "\r\n";
            string individual1Text = "Individual Respondent 1: " + RespondentIndividualOne.Name + "\r\n";
            string individual2Text = "Individual Respondent 2: " + RespondentIndividualTwo.Name;
            string[] linesArray = new String[5] { clientText, company1Text, company2Text, individual1Text, individual2Text };

            File.WriteAllLines(clientInfoPath, linesArray);
        }

        private void PopulateBookmarkList()
        {
            //populate bookmark dictionary
            BookmarkList.Add("ComplainantName", Complainant.Name);
            BookmarkList.Add("ComplainantAddressLine1", Complainant.Address.Line1);
            BookmarkList.Add("ComplainantAddressLine2", Complainant.Address.Line2);
            BookmarkList.Add("DateOfHire", Complainant.DateOfHire.ToString("MMMM d, yyyy"));
            BookmarkList.Add("DateOfTermination", Complainant.DateOfTermination.ToString("MMMM d, yyyy"));
            BookmarkList.Add("OSHARegion", OSHA.RegionNumber);
            BookmarkList.Add("OSHAAddLine1", OSHA.Address.Line1);
            BookmarkList.Add("OSHAAddLine2", OSHA.Address.Line2);
            BookmarkList.Add("OSHAFax", OSHA.FaxNumber);
            BookmarkList.Add("RespondentCompany1Name", RespondentCompanyOne.Name);
            BookmarkList.Add("RespondentCompany1AddressLine1", RespondentCompanyOne.Address.Line1);
            BookmarkList.Add("RespondentCompany1AddressLine2", RespondentCompanyOne.Address.Line2);
            BookmarkList.Add("RespondentCompany2Name", RespondentCompanyTwo.Name);
            BookmarkList.Add("RespondentCompany2AddressLine1", RespondentCompanyTwo.Address.Line1);
            BookmarkList.Add("RespondentCompany2AddressLine2", RespondentCompanyTwo.Address.Line2);
            BookmarkList.Add("RespondentIndividual1Name", RespondentIndividualOne.Name);
            BookmarkList.Add("RespondentIndividual2Name", RespondentIndividualTwo.Name);
            BookmarkList.Add("SigningAttorney", SigningAttorney.Name);
        }

        private void PopulateClientTemplate(KeyValuePair<string, string> templateNameAndfileSaveName)
        {
            toolStripStatusLabel1.Text = "Creating " + templateNameAndfileSaveName.Value;

            object oMissing = Missing.Value;
            string currentDirectory = Directory.GetCurrentDirectory();
            object oPathToTemplate = currentDirectory + "\\" + templateNameAndfileSaveName.Key;
            string absolutePathToSave = Path.GetFullPath(FullPathToClientDirectory);

            if (!File.Exists(oPathToTemplate.ToString()))
            {
                MessageBox.Show($"Cannot find template: {templateNameAndfileSaveName.Key}.\nPlease fill out remaining templates manually.", "Error", MessageBoxButtons.OKCancel);
            }

            Word._Application wordApplication = new Word.Application();
            Word._Document currentDocument = wordApplication.Documents.Add(ref oPathToTemplate, ref oMissing, ref oMissing, ref oMissing);
            wordApplication.Visible = false;

            ReplaceBookmarksWithFormTextIn(currentDocument);

            currentDocument.Repaginate();
            currentDocument.Fields.Update();
            currentDocument.SaveAs2(Path.Combine(absolutePathToSave, templateNameAndfileSaveName.Value));
            currentDocument.Close();
            wordApplication.Quit();
            
        }

        private void ReplaceBookmarksWithFormTextIn(Word._Document document)
        {
            foreach (KeyValuePair<string, string> bookmark in BookmarkList)
            {
                //check for bookmark names, replace with user input text if bookmark exists in template
                if (document.Bookmarks.Exists(bookmark.Key))
                {
                    Word.Bookmark activeBookmark = document.Bookmarks[bookmark.Key];
                    Word.Range range = activeBookmark.Range;

                    range.Text = bookmark.Value;

                    object newRange = range;

                    //reinsert bookmark
                    document.Bookmarks.Add(bookmark.Key, ref newRange);
                }
            }
        }

        private bool IsMissing(string inputValue)
        {
            return inputValue.Length < 2;
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
            Application.Exit();
        }
    }
}