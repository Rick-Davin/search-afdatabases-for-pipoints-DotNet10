namespace Search_AFDatabases_for_PIPoints.Logic
{
    public static class AFOperation
    {
        public static PIServer? ConnectToPIServer(string dataArchiveName)
        {
            PIServer? server = PIServer.FindPIServer(dataArchiveName);
            if (server != null && !server.ConnectionInfo.IsConnected)
            {
                server.Connect();
            }
            return server;
        }

        public static PIServer? ConnectToPIServer(Guid id)
        {
            PIServer? server = PIServer.FindPIServer(id);
            if (server != null && !server.ConnectionInfo.IsConnected)
            {
                server.Connect();
            }
            return server;
        }

        public static PISystem? ConnectToAssetServer(string assetServerName)
        {
            PISystem? server = (new PISystems())[assetServerName];
            if (server != null && !server.ConnectionInfo.IsConnected)
            { 
               server.Connect();
            }
            return server;
        }

        public static IEnumerable<AFDatabase> GetMatchingDatabaseByNameOrPattern(this PISystem assetServer, string databasePattern)
        {
            if (string.IsNullOrWhiteSpace(databasePattern))
            {
                yield break;
            }

            // If there are no wildcards, then the pattern is considered to be an exact name.
            // We return that one name BUT we return as a List item.
            if (!databasePattern.Contains('*') && !databasePattern.Contains('?'))
            {
                yield return assetServer.Databases[databasePattern];
                yield break;
            }

            string pattern = databasePattern.ToUpper();
            foreach (var database in assetServer.Databases.Where(db => IsMatch(db.Name.ToUpper(), pattern)))
            {
                yield return database;
            }
        }

        // Function that matches input str with given wildcard pattern
        // https://www.geeksforgeeks.org/wildcard-pattern-matching/
        private static bool IsMatch(string str, string pattern)
        {

            // lookup table for storing results of
            // subproblems
            bool[] prev = new bool[str.Length + 1];
            bool[] curr = new bool[str.Length + 1];

            // empty pattern can match with empty string
            prev[0] = true;

            // fill the table in bottom-up fashion
            for (int i = 1; i <= pattern.Length; i++)
            {
                bool flag = true;
                for (int ii = 1; ii < i; ii++)
                {
                    if (pattern[ii - 1] != '*')
                    {
                        flag = false;
                        break;
                    }
                }
                curr[0] = flag; // for every row we are assigning
                                // 0th column value.
                for (int j = 1; j <= str.Length; j++)
                {

                    // Two cases if we see a '*'
                    // a) We ignore ‘*’ character and move
                    //    to next character in the pattern,
                    //     i.e., ‘*’ indicates an empty sequence.
                    // b) '*' character matches with ith
                    //     character in input
                    if (pattern[i - 1] == '*')
                    {
                        curr[j] = curr[j - 1] || prev[j];
                    }

                    // Current characters are considered as
                    // matching in two cases
                    // (a) current character of pattern is '?'
                    // (b) characters actually match
                    else if (pattern[i - 1] == '?'
                        || str[j - 1] == pattern[i - 1])
                    {
                        curr[j] = prev[j - 1];
                    }

                    // If characters don't match
                    else
                    {
                        curr[j] = false;
                    }
                }
                prev = (bool[])curr.Clone();
            }

            return prev[str.Length];
        }

    }
}
