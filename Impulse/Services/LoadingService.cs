namespace Impulse.Services
{
    public class LoadingService
    {
        public event Action<bool, string>? OnLoadingStateChanged;
        public bool IsLoading { get; private set; }

        public void Show(string message = "Generating Report...")
        {
            IsLoading = true;
            OnLoadingStateChanged?.Invoke(true, message);
        }
        public void Hide()
        {
            IsLoading = false;
            OnLoadingStateChanged?.Invoke(false, string.Empty);
        }
    }
}
