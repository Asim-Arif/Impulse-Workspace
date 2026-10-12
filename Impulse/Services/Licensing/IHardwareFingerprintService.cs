namespace Impulse.Services.Licensing
{
    public interface IHardwareFingerprintService
    {
        string GetHardwareId();
        bool MatchesHardwareId(string targetHardwareId);
    }
}
