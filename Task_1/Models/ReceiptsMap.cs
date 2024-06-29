using CsvHelper.Configuration;
using Task_1.Models;
public class ReceiptsMap : ClassMap<Receipts>
{
    public ReceiptsMap()
    {
        Map(m => m.businessUnit).Name("Business Unit");
        Map(m => m.receiptMethodId).Name("Receipt Method ID");
        Map(m => m.remittanceBank).Name("Remittance Bank");
        Map(m => m.remittanceBankAccount).Name("Remittance Bank Account");
        Map(m => m.receiptNumber).Name("Receipt Number");
        Map(m => m.receiptAmount).Name("Receipt Amount");
        Map(m => m.receiptDate).Name("Receipt Date");
        Map(m => m.accountingDate).Name("Accounting Date");
        Map(m => m.conversionDate).Name("Conversion Date");
        Map(m => m.currency).Name("Currency");
        Map(m => m.conversionRateType).Name("Conversion Rate Type");
        Map(m => m.conversionRate).Name("Conversion Rate");
        Map(m => m.customerName).Name("Customer Name");
        Map(m => m.customerAccountNumber).Name("Customer Account Number");
        Map(m => m.customerSiteNumber).Name("Customer Site Number");
        Map(m => m.invoiceNumberReference).Name("Invoice number reference");
        Map(m => m.invoiceAmount).Name("Invoice Amount");
        Map(m => m.comments).Name("Comments");
    }
}
