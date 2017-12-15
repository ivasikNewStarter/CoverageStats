using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CoverageStats
{
    class Splitter
    {

        public static String splitEndLineToInClauseOracle(String textToSplit, String field, bool isNumber)
        {
            List<String> splitted = new List<String>();
            long numMatch = 1;
            StringBuilder resultString = new StringBuilder();
            MatchCollection resultMatches = Regex.Matches(textToSplit, @"((?!\r)(?!\n)(?!\s).)+");

            foreach (Match m in resultMatches)
            {
                if (numMatch == 1)
                {
                    resultString.Append(field + " in (");
                }

                if (isNumber)
                {
                    resultString.Append(m.Value);
                }
                else
                {
                    resultString.Append("'" + m.Value + "'");
                }

                if (numMatch % 950 == 0)
                {
                    resultString.Append(")\n or " + field + " in (");
                }


                numMatch++;

                if (numMatch > resultMatches.Count)
                {
                    resultString.Append(")");
                }
                else
                {
                    if ((numMatch - 1) % 950 != 0)
                    {
                        resultString.Append(",");
                    }
                }
            }
            return resultString.ToString();

        }

    }
}
