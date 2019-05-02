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

        private string FullPath;

        private Dictionary<string, string> BookmarkList = new Dictionary<string, string>();

        
        public NewClientIntakeForm()
        {
            InitializeComponent();
        }

        private void TableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void BtnSubmit_Click(object sender, EventArgs e)
        {
            Complainant = new Complainant(
                (isMissing(CompFirstName.Text) && isMissing(CompLastName.Text)) ? "COMPLAINANT NAME" : (CompFirstName.Text.Trim() + " " + CompLastName.Text.Trim()),
                new Address(
                    isMissing(CompAddLine1.Text) ? "COMPLAINANT ADDRESS LINE 1" : CompAddLine1.Text,
                    isMissing(CompAddLine2.Text) ? "COMPLAINANT ADDRESS LINE 2" : CompAddLine2.Text
                    ), 
                DateOfHirePicker.Value, 
                DateOfTerminationPicker.Value, 
                isMissing(CompEmail.Text) ? "COMPLAINANT EMAIL" : CompEmail.Text, 
                isMissing(CompPhone.Text) ? "COMPLAINANT PHONE" : CompPhone.Text
                );

            RespondentCompanyOne = new NamedEntityWithAddress(
                isMissing(RespComp1Name.Text) ? "RESPONDENT COMPANY 1" : RespComp1Name.Text, 
                new Address(
                    isMissing(RespComp1AddLine1.Text) ? "RESPONDENT 1 ADDRESS LINE 1" : RespComp1AddLine1.Text,
                    isMissing(RespComp1AddLine2.Text) ? "RESPONDENT 1 ADDRESS LINE 2" : RespComp1AddLine2.Text
                    )
                );

            RespondentCompanyTwo = new NamedEntityWithAddress(
                isMissing(RespComp2Name.Text) ? "RESPONDENT COMPANY 2" : RespComp2Name.Text,
                new Address(
                    isMissing(RespComp2AddLine1.Text) ? "RESPONDENT 2 ADDRESS LINE 1" : RespComp2AddLine1.Text,
                    isMissing(RespComp2AddLine2.Text) ? "RESPONDENT 2 ADDRESS LINE 2" : RespComp2AddLine2.Text
                    )
                );

            RespondentIndividualOne = new NamedEntity(isMissing(RespInd1Name.Text) ? "RESPONDENT INDIVIDUAL 1" : RespInd1Name.Text);

            RespondentIndividualTwo = new NamedEntity(isMissing(RespInd2Name.Text) ? "RESPONDENT INDIVIDUAL 2" : RespInd2Name.Text);

            SigningAttorney = new NamedEntity(isMissing(AttorneyName.Text) ? "SIGNING ATTORNEY" : AttorneyName.Text);

            OSHA = new OSHARegion(selectOSHA.Text);

            //create client folder structure
            string basepath = "..\\..\\Active Clients\\OSHA";
            string clientFolder = $"{CompLastName.Text.Trim()}, {CompFirstName.Text.Trim()}";
            FullPath = basepath + "\\" + clientFolder;

            //create client folder in OSHA folder
            Directory.CreateDirectory(FullPath);

            //create case information text document
            string clientInfoPath = FullPath + "\\Client Information.txt";
            string clientText = "Complainant\r\n--------------------\r\nName: " + Complainant.Name + "\r\nAddress:\r\n" + Complainant.Address.Line1 + "\r\n" + Complainant.Address.Line2 + "\r\nPhone: " + Complainant.PhoneNumber + "\r\nEmail: " + Complainant.EmailAddress + "\r\nDate Of Hire: " + Complainant.DateOfHire.ToString("D") + "\r\nDate Of Termination: " + Complainant.DateOfTermination.ToString("D") + "\r\n";
            string company1Text = "Respondent Company 1\r\n--------------------\r\nName: " + RespondentCompanyOne.Name + "\r\nAddress:\r\n" + RespondentCompanyOne.Address.Line1 + "\r\n" + RespondentCompanyOne.Address.Line2 + "\r\n";
            string company2Text = "Respondent Company 2\r\n--------------------\r\nName: " + RespondentCompanyTwo.Name + "\r\nAddress:\r\n" + RespondentCompanyTwo.Address.Line1 + "\r\n" + RespondentCompanyTwo.Address.Line2 + "\r\n";
            string individual1Text = "Individual Respondent 1: " + RespondentIndividualOne.Name + "\r\n";
            string individual2Text = "Individual Respondent 2: " + RespondentIndividualTwo.Name;
            string[] linesArray = new String[5] { clientText, company1Text, company2Text, individual1Text, individual2Text };

            File.WriteAllLines(clientInfoPath, linesArray);

            //populate bookmark dictionary
            BookmarkList.Add("ComplainantName", Complainant.Name);
            BookmarkList.Add("ComplainantAddressLine1", Complainant.Address.Line1);
            BookmarkList.Add("ComplainantAddressLine2", Complainant.Address.Line2);
            BookmarkList.Add("DateOfHire", Complainant.DateOfHire.ToString("D"));
            BookmarkList.Add("DateOfTermination", Complainant.DateOfTermination.ToString("D"));
            BookmarkList.Add("OSHARegion", OSHA.RegionNumber);
            BookmarkList.Add("OSHAAddLine1", OSHA.Address.Line1);
            BookmarkList.Add("OSHAAddLine2", OSHA.Address.Line2);
            BookmarkList.Add("OSHAFax", OSHA.FaxNumber);
            BookmarkList.Add("RespondentCompany1Name", RespondentCompanyOne.Name);
            BookmarkList.Add("RespondentCompany1AddressLine1",RespondentCompanyOne.Address.Line1);
            BookmarkList.Add("RespondentCompany1AddressLine2",RespondentCompanyOne.Address.Line2);
            BookmarkList.Add("RespondentCompany2Name",RespondentCompanyTwo.Name);
            BookmarkList.Add("RespondentCompany2AddressLine1",RespondentCompanyTwo.Address.Line1);
            BookmarkList.Add("RespondentCompany2AddressLine2",RespondentCompanyTwo.Address.Line2);
            BookmarkList.Add("RespondentIndividual1Name",RespondentIndividualOne.Name);
            BookmarkList.Add("RespondentIndividual2Name",RespondentIndividualTwo.Name);
            BookmarkList.Add("SigningAttorney", SigningAttorney.Name);

            //create complaint template
            PopulateClientTemplate("New Complaint Template.dotx", "Complaint Draft v1.docx", BookmarkList);

            //create anatomy letter
            PopulateClientTemplate("Anatomy Letter Template.dotx", "Anatomy Letter Draft v1.docx", BookmarkList);

            //create retainer
            PopulateClientTemplate("Retainer Template.dotx", "Retainer Draft v1.docx", BookmarkList);
            
            //close this form
            this.Close();
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void PopulateClientTemplate(string templateName, string fileSaveName, Dictionary<string, string> bookmarks)
        {
            object oMissing = Missing.Value;
            string currentDirectory = Directory.GetCurrentDirectory();
            object oPathToTemplate = currentDirectory + "\\" + templateName;
            string absolutePathToSave = Path.GetFullPath(FullPath);

            if (!File.Exists(oPathToTemplate.ToString()))
            {
                MessageBox.Show($"Cannot find template: {templateName}.\nPlease fill out remaining templates manually.", "Error", MessageBoxButtons.OKCancel);
            }

            Word._Application wordApplication = new Word.Application();
            Word._Document document = wordApplication.Documents.Add(ref oPathToTemplate, ref oMissing, ref oMissing, ref oMissing);
            wordApplication.Visible = false;


            foreach(KeyValuePair<string, string> bookmark in bookmarks )
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

            document.Repaginate();
            document.Fields.Update();
            document.SaveAs2(Path.Combine(absolutePathToSave, fileSaveName));
            document.Close();
            wordApplication.Quit();
        }

        private bool isMissing(string inputValue)
        {
            return inputValue.Length < 2;
        }
    }
}