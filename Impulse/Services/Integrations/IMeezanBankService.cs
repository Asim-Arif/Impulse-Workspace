namespace CIP.Services
{
    public interface IMeezanBankService
    {
        Task<string> FetchTitleMeezan(string strDebitAccount, string strCreditAccount, string strTransactionType, string strBankCode = "");
        Task<bool> PostToMeezanPortal(int entryID, int transactionType);
    }
}
