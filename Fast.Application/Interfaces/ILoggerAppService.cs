namespace Fast.Application.Interfaces
{
	public interface ILoggerAppService 
	{		
		void LogInfo(string message, long userID = 0, string userName = null);
		void LogWarning(string message, long userID = 0, string userName = null);
		void LogDebug(string message, long userID = 0, string userName = null);
		void LogError(string message, long userID = 0, string userName = null);
	}
}