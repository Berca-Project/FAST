namespace Fast.Domain.Interfaces.Services
{
	public interface ILoggerService
	{		
		void LogInfo(string message, long userID, string userName);
		void LogWarning(string message, long userID, string userName);
		void LogDebug(string message, long userID, string userName);
		void LogError(string message, long userID, string userName);
	}
}
