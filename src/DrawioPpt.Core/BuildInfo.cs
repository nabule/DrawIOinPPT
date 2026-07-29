using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace DrawioPpt.Core
{
    public static class BuildInfo
    {
        private const string DefaultProductVersion = "v1.0.8";
        private const string UnknownGitShortHash = "unknown";
        private const string BuildInfoFileName = "BuildInfo.txt";
        private const string PackageFileName = "PACKAGE.txt";
        private static readonly Lazy<Dictionary<string, string>> Values =
            new Lazy<Dictionary<string, string>>(LoadValues);

        public static string ProductVersion
        {
            get { return ReadValue("ProductVersion", ReadValue("Version", DefaultProductVersion)); }
        }

        public static string GitShortHash
        {
            get { return ReadValue("GitShortHash", UnknownGitShortHash); }
        }

        public static string BuildId
        {
            get
            {
                string buildId = ReadValue("BuildId", string.Empty);
                if (!string.IsNullOrWhiteSpace(buildId))
                {
                    return buildId;
                }

                string gitShortHash = GitShortHash;
                return string.Equals(gitShortHash, UnknownGitShortHash, StringComparison.OrdinalIgnoreCase)
                    ? ProductVersion
                    : ProductVersion + "+" + gitShortHash;
            }
        }

        public static string DisplayVersion
        {
            get { return BuildId; }
        }

        private static string ReadValue(string key, string defaultValue)
        {
            string value;
            if (Values.Value.TryGetValue(key, out value) && !string.IsNullOrWhiteSpace(value))
            {
                return value.Trim();
            }

            return defaultValue;
        }

        private static Dictionary<string, string> LoadValues()
        {
            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string path in GetCandidatePaths())
            {
                if (!File.Exists(path))
                {
                    continue;
                }

                foreach (string line in File.ReadAllLines(path))
                {
                    string trimmed = line == null ? string.Empty : line.Trim();
                    if (trimmed.Length == 0 || trimmed.StartsWith("#", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    int separatorIndex = trimmed.IndexOf(':');
                    if (separatorIndex < 0)
                    {
                        separatorIndex = trimmed.IndexOf('=');
                    }

                    if (separatorIndex <= 0)
                    {
                        continue;
                    }

                    string key = trimmed.Substring(0, separatorIndex).Trim();
                    string value = trimmed.Substring(separatorIndex + 1).Trim();
                    if (!values.ContainsKey(key))
                    {
                        values.Add(key, value);
                    }
                }

                if (values.ContainsKey("BuildId") || values.ContainsKey("GitShortHash"))
                {
                    break;
                }
            }

            return values;
        }

        private static IEnumerable<string> GetCandidatePaths()
        {
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            foreach (string path in GetCandidatePathsForDirectory(baseDirectory))
            {
                yield return path;
            }

            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            if (!string.IsNullOrWhiteSpace(assemblyPath))
            {
                foreach (string path in GetCandidatePathsForDirectory(Path.GetDirectoryName(assemblyPath)))
                {
                    yield return path;
                }
            }
        }

        private static IEnumerable<string> GetCandidatePathsForDirectory(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                yield break;
            }

            yield return Path.Combine(directory, BuildInfoFileName);
            yield return Path.Combine(directory, PackageFileName);

            DirectoryInfo parent = Directory.GetParent(directory);
            if (parent != null)
            {
                yield return Path.Combine(parent.FullName, BuildInfoFileName);
                yield return Path.Combine(parent.FullName, PackageFileName);
            }
        }
    }
}
