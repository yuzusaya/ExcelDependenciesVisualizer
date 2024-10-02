using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDependenciesVisualizer.Utilities;

public static class StringUtility
{
    public static List<string> SplitMultiDelims(this string str, string delims)
    {
        var tokens = new HashSet<string>();

        // Replace each delimiter with a space
        foreach (char delim in delims)
        {
            str = str.Replace(delim.ToString(), " ");
        }

        // Split the string based on spaces
        var splitTokens = str.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

        // Add unique tokens to the HashSet
        foreach (var token in splitTokens)
        {
            tokens.Add(token.Trim());
        }

        return new List<string>(tokens);
    }
}
