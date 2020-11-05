using Fast.Domain.Interfaces.Services;
using NLog;
using System;

namespace Fast.Infra.Logging
{
	public class Logger : ILoggerService
	{
		private static readonly ILogger logger = LogManager.GetCurrentClassLogger();

		public void LogDebug(string message, long userID, string userName)
		{
			LogEventInfo theEvent = new LogEventInfo(LogLevel.Debug, "Fast Application", message);
			theEvent.Properties["UserID"] = (int)userID;
            theEvent.Properties["UserName"] = userName;
            logger.Log(theEvent);
		}

		public void LogError(string message, long userID, string userName)
		{
			LogEventInfo theEvent = new LogEventInfo(LogLevel.Error, "Fast Application", message);
			theEvent.Properties["UserID"] = (int)userID;
            theEvent.Properties["UserName"] = userName;
            logger.Log(theEvent);
		}

		public void LogInfo(string message, long userID, string userName)
		{
			LogEventInfo theEvent = new LogEventInfo(LogLevel.Info, "Fast Application", message);
			theEvent.Properties["UserID"] = (int)userID;
            theEvent.Properties["UserName"] = userName;
            logger.Log(theEvent);
		}

		public void LogWarning(string message, long userID, string userName)
		{
			LogEventInfo theEvent = new LogEventInfo(LogLevel.Warn, "Fast Application", message);
			theEvent.Properties["UserID"] = (int)userID;
            theEvent.Properties["UserName"] = userName;
            logger.Log(theEvent);
		}
	}
}
