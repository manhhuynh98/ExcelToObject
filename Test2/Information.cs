using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Test2
{
    public class Information
    {
        public string fullName { get; set; }
        public string address1 { get; set; }
        public string address2 { get; set; }
        public string city { get; set; }
        public string stateCode { get; set; }
        public string zipCode { get; set; }
        public string countryName { get; set; }
        public string productName { get; set; }
        public string productType { get; set; }
        public string productColor { get; set; }
        public string productSize { get; set; }
        public string quantity { get; set; }
        public string productID { get; set; }

        public Information()
        {
        }

        public Information(string fullName, string address1, string address2, string city, string stateCode, string zipCode, string countryName, string productName, string productType, string productColor, string productSize, string quantity, string productID)
        {
            this.fullName = fullName;
            this.address1 = address1;
            this.address2 = address2;
            this.city = city;
            this.stateCode = stateCode;
            this.zipCode = zipCode;
            this.countryName = countryName;
            this.productName = productName;
            this.productType = productType;
            this.productColor = productColor;
            this.productSize = productSize;
            this.quantity = quantity;
            this.productID = productID;
        }
    }
}
