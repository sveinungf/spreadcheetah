using System;

namespace SpreadCheetah
{
    public class SpreadCheetahException : Exception
    {
        public SpreadCheetahException()
        {
        }

        public SpreadCheetahException(string message) : base(message)
        {
        }

        public SpreadCheetahException(string message, Exception exception) : base(message, exception)
        {
        }
    }
}
