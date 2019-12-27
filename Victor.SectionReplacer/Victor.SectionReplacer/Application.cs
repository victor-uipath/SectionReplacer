
namespace Victor.SectionReplacer
{
    /// <summary>
    /// The Application class represents the starting point for the backend work using your chosen SDK.
    /// </summary>
    public class Application
    {

        #region Constructors

        // Creates a new, blank Application 
        public Application() { }

        //Make the sum of 2 values
        public int Sum(int firstValue, int secondValue)
        {
            return firstValue + secondValue;
        }
        // concatenate strings
        public string Concatenate(string a, string b, string c, string d)
        {
            return a + " " + b + " " + c + " " + d;
        }

        #endregion
    }
}
