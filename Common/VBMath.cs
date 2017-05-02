using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Common
{
    /// <summary>
    /// 
    /// 1. C# does not support integer division
    /// 2. VB6 automatic convert to double and round after result returned in case divide, multiply two number and store in whole number
    /// </summary>
    public static class VBMath
    {
        /// <summary>
        /// Support \
        /// </summary>
        /// <param name="devidend"></param>
        /// <param name="divisor"></param>
        /// <returns>Result of integer division</returns>
        public static int IntegerDivision(double devidend, double divisor)
        {
            return (int)Math.Round(devidend) / (int)Math.Round(divisor);
        }

        /// <summary>
        /// Implicit round when devide and store value in integer variable
        /// For another type like short, byte,.. just using normal casting
        /// </summary>
        /// <param name="devidend"></param>
        /// <param name="divisor"></param>
        /// <returns></returns>
        public static int Devision(double devidend, double divisor)
        {
            return (int) Math.Round(devidend / divisor);
        }

        /// <summary>
        /// Implicit multiply and round.
        /// </summary>
        /// <param name="multiplier"></param>
        /// <param name="multiplicand"></param>
        /// <returns></returns>
        public static int Multiplication(double multiplier, double multiplicand)
        {
            return (int) Math.Round(multiplier * multiplicand);
        }
    }
}
