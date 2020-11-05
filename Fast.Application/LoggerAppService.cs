using Fast.Application.Interfaces;
using Fast.Domain.Interfaces.Services;

namespace Fast.Application
{
	public class LoggerAppService : ILoggerAppService
	{
		public readonly ILoggerService _service;

		public LoggerAppService(ILoggerService serviceBase)
		{
			_service = serviceBase;
		}

		public void LogDebug(string message, long userID = 0, string userName = null)
		{
			_service.LogDebug(message, userID, userName);
		}

		public void LogError(string message, long userID = 0, string userName = null)
		{
			_service.LogError(message, userID, userName);
		}

		public void LogInfo(string message, long userID = 0, string userName = null)
		{
			_service.LogInfo(message, userID, userName);
		}

		public void LogWarning(string message, long userID = 0, string userName = null)
		{
			_service.LogWarning(message, userID, userName);
		}
	}
}