namespace Fast.Infra.CrossCutting.Common
{
	public class QueryFilter
	{
		public QueryFilter(string propertyName, int value)
		{
			PropertyName = propertyName;
			Value = value.ToString();
		}

        public QueryFilter(string propertyName, long value)
        {
            PropertyName = propertyName;
            Value = value.ToString();
        }

        public QueryFilter(string propertyName, string value)
		{
			PropertyName = propertyName;
			Value = value;
		}

		public QueryFilter(string propertyName, string value, Operator operatorValue)
		{
			PropertyName = propertyName;
			Value = value;
			Operator = operatorValue;
		}

		public QueryFilter(string propertyName, string value, Operator operatorValue, Operation operationValue)
		{
			PropertyName = propertyName;
			Value = value;
			Operator = operatorValue;
			Operation = operationValue;
		}

		public string PropertyName { get; set; }
		public string Value { get; set; }
		public Operator Operator { get; set; } = Operator.Equals;

		public Operation Operation { get; set; } = Operation.AndAlso;
	}
}
