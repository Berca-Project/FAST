using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web.Mvc;

namespace Fast.Web.Utils
{
	public static class ExceptionExtensions
	{
		public static string GetModelStateErrors(this ModelStateDictionary state)
		{
			string messages = string.Join("; ", state.Values
										.SelectMany(x => x.Errors)
										.Select(x => x.ErrorMessage));

			return messages;
		}

		public static IEnumerable<Exception> GetAllExceptions(this Exception ex)
		{
			Exception currentEx = ex;
			yield return currentEx;
			while (currentEx.InnerException != null)
			{
				currentEx = currentEx.InnerException;
				yield return currentEx;
			}
		}

		public static IEnumerable<string> GetAllExceptionAsString(this Exception ex)
		{
			Exception currentEx = ex;
			yield return currentEx.ToString();
			while (currentEx.InnerException != null)
			{
				currentEx = currentEx.InnerException;
				yield return currentEx.ToString();
			}
		}

		public static IEnumerable<string> GetAllExceptionMessages(this Exception ex)
		{
			Exception currentEx = ex;
			yield return currentEx.Message;
			while (currentEx.InnerException != null)
			{
				currentEx = currentEx.InnerException;
				yield return currentEx.Message;
			}
		}
		public static string GetAllMessages(this Exception ex)
		{
			if (ex == null)
				throw new ArgumentNullException("ex");

			StringBuilder sb = new StringBuilder();

			while (ex != null)
			{
				if (!string.IsNullOrEmpty(ex.Message))
				{
					if (sb.Length > 0)
						sb.Append(Environment.NewLine);

					sb.Append(ex.Message);
				}

				ex = ex.InnerException;
			}

			return sb.ToString();
		}
	}
}