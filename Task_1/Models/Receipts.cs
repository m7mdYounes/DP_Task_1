using Microsoft.VisualBasic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace Task_1.Models
{
    public class Receipts
    {
        [Display(Name = "Business Unit")]
        public string? businessUnit { get; set; }
        [Display(Name = "Receipt Method Id")]
        public string? receiptMethodId { get; set;}
        [Display(Name = "Remittance Bank")]
        public string? remittanceBank { get; set;}
        [Display(Name = "Remittance Bank Account")]
        public string? remittanceBankAccount { get; set; }
        [Display(Name = "Receipt Number")]
        public string? receiptNumber { get; set; }
        [Display(Name = "Receipt Amount")]
        public decimal? receiptAmount { get; set; }
        [Display(Name = "Receipt Date")]
        public string? receiptDate { get; set; }
        [Display(Name = "Accounting Date")]
        public string? accountingDate { get; set; }
        [Display(Name = "Conversion Date")]
        public string? conversionDate { get; set; }
        [Display(Name = "Currency")]
        public string? currency { get; set; }
        [Display(Name = "Conversion Rate Type")]
        public string? conversionRateType { get; set; }
        [Display(Name = "Conversion Rate")]
        public string? conversionRate { get; set; }
        [Display(Name = "Customer Name")]
        public string? customerName { get; set; }
        [Display(Name = "Customer Account Number")]
        public string? customerAccountNumber { get; set; }
        [Display(Name = "Customer Site Number")]
        public string? customerSiteNumber { get; set; }
        [Display(Name = "Invoice Number Reference")]
        public string? invoiceNumberReference { get; set; }
        [Display(Name = "Invoice Amount")]
        public decimal? invoiceAmount { get; set; }
        [Display(Name = "Comments")]
        public string? comments { get; set; }

    }
}
