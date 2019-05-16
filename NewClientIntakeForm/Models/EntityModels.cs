using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NewClientIntakeForm.Models
{
    public class NamedEntity
    {
        public string Name { get; private set; }

        public NamedEntity(string name)
        {
            this.Name = name.Length < 3 ? "NAME" : name.Trim();
        }

        public NamedEntity(string firstName, string lastName)
        {
            if (firstName.Length > 1 && lastName.Length > 1)
            {
                this.Name = firstName.Trim() + " " + lastName.Trim();
            }
            else
            {
                this.Name = "NAME";
            }
            
        }
    }

    public class Address
    {
        public string Line1 { get; private set; } = "ADDRESS LINE 1";
        public string Line2 { get; private set; } = "ADDRESS LINE 2";

        public Address() { }
        public Address(string line1, string line2)
        {
            this.Line1 = line1.Length < 2 ? "ADDRESS LINE 1" : line1;
            this.Line2 = line2.Length < 2 ? "ADDRESS LINE 2" : line2;
        }

        public string GetFullAddressAsString()
        {
            string fullAddress = Line1 + ", " + Line2;
            return fullAddress;
        }
    }

    public class NamedEntityWithAddress : NamedEntity
    {
        public Address Address { get; private set; } = new Address();

        public NamedEntityWithAddress(string name, Address address) : base(name)
        {
            this.Address = address;
        }

        public NamedEntityWithAddress(string firstName, string lastName, Address address) : base(firstName, lastName)
        {
            this.Address = address;
        }
    }

    public class Complainant : NamedEntityWithAddress
    {
        public DateTime DateOfHire { get; private set; } = DateTime.Now;
        public DateTime DateOfTermination { get; private set; } = DateTime.Now;
        public string EmailAddress { get; private set; } = "EMAIL ADDRESS";
        public string PhoneNumber { get; private set; } = "PHONE NUMBER";
        public Complainant(string fullName, Address address, DateTime dateOfHire, DateTime dateOfTermination, string emailAddress, string phoneNumber) : base(fullName, address)
        {
            this.DateOfHire = dateOfHire;
            this.DateOfTermination = dateOfTermination;
            this.EmailAddress = emailAddress.Length < 3 ? "EMAIL ADDRESS" : emailAddress;
            this.PhoneNumber = phoneNumber.Length < 3 ? "PHONE NUMBER" : phoneNumber;
        }
        public Complainant(string firstName, string lastName, Address address, DateTime dateOfHire, DateTime dateOfTermination, string emailAddress, string phoneNumber) : base(firstName, lastName, address)
        {
            this.DateOfHire = dateOfHire;
            this.DateOfTermination = dateOfTermination;
            this.EmailAddress = emailAddress.Length < 3 ? "EMAIL ADDRESS" : emailAddress;
            this.PhoneNumber = phoneNumber.Length < 3 ? "PHONE NUMBER" : phoneNumber;
        }
    }

    public class RespondentCompany : NamedEntityWithAddress
    {
        public RespondentCompany() : base("RESPONDENT COMPANY NAME", new Address("RESPONDENT ADDRESS LINE 1", "RESPONDENT ADDRESS LINE 2"))
        {

        }
        public RespondentCompany(string name, Address address) : base(name, address)
        {

        }
    }

    public class OSHARegion
    {
        public string RegionNumber { get; private set; } = "REGION NUMBER";
        public Address Address { get; private set; } = new Address();
        public string FaxNumber { get; private set; } = "FAX NUMBER";

        public OSHARegion()
        {
            this.RegionNumber = null;
            this.Address = getAddressByRegion("");
            this.FaxNumber = getFaxNumberByRegion("");
        }

        public OSHARegion(string regionNumber)
        {
            this.RegionNumber = regionNumber;
            this.Address = getAddressByRegion(regionNumber);
            this.FaxNumber = getFaxNumberByRegion(regionNumber);
        }

        private Address getAddressByRegion(string regionNumber)
        {
            Address address;

            switch(regionNumber)
            {
                case "I":
                    address = new Address("JFK Federal Building, 25 New Sudbury Street, Room E340", "Boston, Massachusetts 02203");
                    break;
                case "II":
                    address = new Address("Federal Building, 201 Varick Street, Room 670", "New York, New York 10014");
                    break;
                case "III":
                    address = new Address("The Curtis Center-Suite 740 West, 170 S. Independence Mall West", "Philadelphia, PA 19106-3309");
                    break;
                case "IV":
                    address = new Address("Sam Nunn Atlanta Federal Center, 61 Forsyth Street, SW, Room 6T50", "Atlanta, Georgia 30303");
                    break;
                case "V":
                    address = new Address("John C. Kluczynski Federal Building, 230 South Dearborn Street, Room 3244", "Chicago, Illinois 60604");
                    break;
                case "VI":
                    address = new Address("A. Maceo Smith Federal Building, 525 Griffin Street, Suite 602", "Dallas, Texas 75202");
                    break;
                case "VII":
                    address = new Address("Two Pershing Square Building, 2300 Main Street, Suite 1010", "Kansas City, Missouri 64108-2416");
                    break;
                case "VIII":
                    address = new Address("Cesar Chavez Memorial Building, 1244 Speer Blvd., Suite 551", "Denver, CO 80204");
                    break;
                case "IX":
                    address = new Address("San Francisco Federal Building, 90 7th Street, Suite 2650", "San Francisco, California 94103");
                    break;
                case "X":
                    address = new Address("Fifth & Yesler Tower, 300 Fifth Avenue, Suite 1280", "Seattle, Washington 98104");
                    break;
                default:
                    address = new Address("OSHA ADDRESS LINE ONE", "OSHA ADDRESS LINE TWO");
                    break;
            }
            return address;
        }
        private string getFaxNumberByRegion(string regionNumber)
        {
            string faxNumber;

            switch (regionNumber)
            {
                case "I":
                    faxNumber = "(617) 565-9827";
                    break;
                case "II":
                    faxNumber = "(212) 337-2371";
                    break;
                case "III":
                    faxNumber = "(215) 861-4904";
                    break;
                case "IV":
                    faxNumber = "(678) 237-0447";
                    break;
                case "V":
                    faxNumber = "(312) 353-7774";
                    break;
                case "VI":
                    faxNumber = "(972) 850-4149";
                    break;
                case "VII":
                    faxNumber = "(816) 283-0547";
                    break;
                case "VIII":
                    faxNumber = "720-264-6585";
                    break;
                case "IX":
                    faxNumber = "(415) 625-2534";
                    break;
                case "X":
                    faxNumber = "(206) 757-6705";
                    break;
                default:
                    faxNumber = "OSHA FAX NUMBER";
                    break;
            }
            return faxNumber;
        }
    }
}