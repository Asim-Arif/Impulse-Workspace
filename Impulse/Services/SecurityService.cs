namespace Impulse.Services
{
    public class SecurityService
    {
        // This event tells MainLayout to show the modal
        public event Func<string, Task<bool>> OnPasswordRequest;

        public async Task<bool> VerifyActionAsync(string settingName)
        {
            if (OnPasswordRequest != null)
            {
                return await OnPasswordRequest.Invoke(settingName);
            }
            return false;
        }
    }
}
