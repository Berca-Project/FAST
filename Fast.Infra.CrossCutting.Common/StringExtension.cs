using System;
using System.Text;

namespace Fast.Infra.CrossCutting.Common
{
	public static class StringExtension
	{
		public static string GetExceptionMessages(this Exception ex)
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
