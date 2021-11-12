namespace Microsoft.Azure
{
    public class GraphFilterExpression
    {
        public GraphFilterExpression(string filterProperty, string filterOperator, string filterValue)
        {
            FilterOperator = filterOperator;
            FilterProperty = filterProperty;
            FilterValue = filterValue;            
        }
        public string FilterProperty {get;}
        public string FilterOperator{get;}
        public string FilterValue {get;}

        public override string ToString()
        {
            return $"{FilterProperty} {FilterOperator} '{FilterValue}'";
        }
    }
}