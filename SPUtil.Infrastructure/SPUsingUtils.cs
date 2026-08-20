using Microsoft.Win32;

using System.Net;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Serilog;

namespace SPUtil.Infrastructure
{
    public static class SPUsingUtils
    {
        private static readonly ILogger _log = Log.ForContext("SourceContext", nameof(SPUsingUtils));

		
        private static SecureString DecryptFromPowerShell(string hexString)
		{
			if (string.IsNullOrEmpty(hexString)) return new SecureString();
			byte[] encryptedBytes = Enumerable.Range(0, hexString.Length / 2)
				.Select(x => Convert.ToByte(hexString.Substring(x * 2, 2), 16)).ToArray();
			byte[] decryptedBytes = ProtectedData.Unprotect(encryptedBytes, null, DataProtectionScope.CurrentUser);
			string plainText = Encoding.Unicode.GetString(decryptedBytes);
			var secureString = new SecureString();
			foreach (char c in plainText) secureString.AppendChar(c);
			secureString.MakeReadOnly();
			return secureString;
		}

        /// <summary>
        /// Splits a user name into its domain prefix and the account name.
        /// Accepted input is either "user" or "DOMAIN\user" — anything else is rejected
        /// by the caller. The returned domain is empty when no prefix was given.
        /// </summary>
        public static (string Domain, string User) SplitUserName(string input)
        {
            string value = (input ?? string.Empty).Trim();
            if (value.Length == 0) return (string.Empty, string.Empty);

            int slash = value.IndexOf('\\');
            return slash < 0
                ? (string.Empty, value)
                : (value.Substring(0, slash).Trim(), value.Substring(slash + 1).Trim());
        }	
        public static string NormalizeUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return url;

            try
            {
                Uri uri = new Uri(url);
                string host = uri.Host;
                string[] parts = host.Split('.');

                if (parts.Length > 0)
                {
                    string firstPart = parts[0];

                    if (firstPart.EndsWith("2"))
                    {
                        parts[0] = firstPart.Remove(firstPart.Length - 1);

                        var builder = new UriBuilder(uri);
                        builder.Host = string.Join(".", parts);
                        return builder.Uri.ToString().TrimEnd('/');
                    }
                }
            }
            catch
            {
                // If URL is invalid, return as-is
            }

            return url.Trim();
        }


        /// <summary>
        /// Removes every stored credential profile. Only the Profiles subtree is
        /// deleted — values kept directly under CrSiteAutomate belong to other tooling
        /// and are left untouched.
        /// Returns the number of profiles that were removed.
        /// </summary>
        public static int ForgetAllProfiles()
        {
            string profilesPath = $@"{regPath}\Profiles";

            using (var key = Registry.CurrentUser.OpenSubKey(profilesPath))
            {
                if (key == null)
                {
                    _log.Debug("ForgetAllProfiles — nothing to remove");
                    return 0;
                }

                int count = key.SubKeyCount;
                _log.Information("ForgetAllProfiles — removing {Count} profile(s): {Names}",
                    count, string.Join(", ", key.GetSubKeyNames()));

                // The key must be closed before the tree can be deleted
                key.Close();
                Registry.CurrentUser.DeleteSubKeyTree(profilesPath, throwOnMissingSubKey: false);

                return count;
            }
        }

        /// <summary>Names of the domains that currently have a stored profile.</summary>
        public static string[] GetProfileDomains()
        {
            using (var key = Registry.CurrentUser.OpenSubKey($@"{regPath}\Profiles"))
                return key?.GetSubKeyNames() ?? Array.Empty<string>();
        }		
		
        public static string UrlWithF5(string url)
        {
            if (string.IsNullOrWhiteSpace(url)) return url;

            try
            {
                var uri = new Uri(url);
                string host = uri.Host; // например, portals.ekmd.huji.ac.il
                string[] parts = host.Split('.');

                if (parts.Length > 0)
                {
                    string firstPart = parts[0];
                    // Если первый сегмент не заканчивается на '2', добавляем её
                    if (!firstPart.EndsWith("2"))
                    {
                        parts[0] = firstPart + "2";
                        string newHost = string.Join(".", parts);
                        return url.Replace(host, newHost);
                    }
                }
            }
            catch (Exception ex)
            {
                _log.Caller().Error(ex, "UrlWithF5: failed to parse/normalize URL — {Url}", url);
            }

            return url;
        }

        public static string GetCleanFieldXml(string rawXml, bool forComparison = false)
        {
            try
            {
                var xDoc = System.Xml.Linq.XDocument.Parse(rawXml);
                var root = xDoc.Root;
                if (root == null) return rawXml;

                // Список атрибутов, которые мы ОСТАВЛЯЕМ (все остальное — системный мусор)
                string[] attributesToKeep =
                {
            "Name", "Type", "DisplayName", "Required", "Format",
            "ShowField", "List", "Formula", "ResultType",
            "MaxLength", "Choices", "Default", "StaticName", "Mult"
        };

                string fieldType = root.Attribute("Type")?.Value ?? "";

                // Обработка атрибутов
                var query = root.Attributes()
                    .Where(attr => attributesToKeep.Contains(attr.Name.LocalName));

                if (forComparison)
                {
                    // НОРМАЛИЗАЦИЯ для сравнения
                    query = query.Select(attr => {
                        // Игнорируем стандартный MaxLength для текста (дефолт в SP)
                        if (fieldType == "Text" && attr.Name.LocalName == "MaxLength" && attr.Value == "255")
                            return null;

                        // Можно добавить: если Required="FALSE", тоже считаем дефолтом
                        // if (attr.Name.LocalName == "Required" && attr.Value.ToUpper() == "FALSE") return null;

                        return attr;
                    }).OfType<XAttribute>(); //.Where(attr => attr != null);

                    // Сортируем для стабильного сравнения строк
                    query = query.OrderBy(attr => attr.Name.LocalName);
                }

                XElement cleanField = new XElement("Field", query);

                // Обработка вложенных элементов
                foreach (var child in root.Elements())
                {
                    if (forComparison && child.Name.LocalName == "Choices")
                    {
                        // Сортируем варианты выбора, чтобы разный порядок не считался ошибкой
                        var sortedChoices = new XElement("Choices",
                            child.Elements().OrderBy(e => e.Value));
                        cleanField.Add(sortedChoices);
                    }
                    else
                    {
                        cleanField.Add(new XElement(child));
                    }
                }

                return cleanField.ToString();
            }
            catch
            {
                return rawXml;
            }
        }
        public static string FormatXml(string xml)
        {
            if (string.IsNullOrWhiteSpace(xml)) return xml;

            try
            {
                // Если в строке несколько элементов (как у нас), 
                // оборачиваем их во временный корень <Root>
                string wrappedXml = $"<Root>{xml}</Root>";
                var xDoc = XDocument.Parse(wrappedXml);

                var settings = new System.Xml.XmlWriterSettings
                {
                    Indent = true,
                    IndentChars = "  ",
                    NewLineChars = Environment.NewLine,
                    OmitXmlDeclaration = true,
                    ConformanceLevel = System.Xml.ConformanceLevel.Fragment // Позволяет работать с фрагментами
                };

                using (var stringWriter = new System.IO.StringWriter())
                {
                    using (var xmlWriter = System.Xml.XmlWriter.Create(stringWriter, settings))
                    {
                        // Пишем только содержимое нашего виртуального корня
                        foreach (var node in xDoc.Root!.Nodes())
                        {
                            node.WriteTo(xmlWriter);
                        }
                    }
                    return stringWriter.ToString();
                }
            }
            catch (Exception ex)
            {
                _log.Caller().Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                // Если даже так не вышло, возвращаем оригинал, чтобы не падать
                System.Diagnostics.Debug.WriteLine("XML formatting error: " + ex.Message);
                return xml;
            }
        }
		

        private static string regPath = @"SOFTWARE\Microsoft\CrSiteAutomate";

        /// <summary>
        /// Registry path of the credential profile for a given SharePoint site.
        /// Layout matches the PowerShell tooling so both share the same profiles:
        ///   HKCU\SOFTWARE\Microsoft\CrSiteAutomate\Profiles\&lt;domain&gt;
        /// </summary>
        private static string ProfilePath(string domain) => $@"{regPath}\Profiles\{domain}";

        /// <summary>
        /// Resolves the AD domain from the site FQDN — the second host segment.
        ///   https://crs.ada.huji.ac.il/...   → "ada"
        ///   https://tss2.ekmd.huji.ac.il/... → "ekmd"
        /// The two farms live in separate domains with no trust between them, so the
        /// domain cannot be taken from the current Windows session: that would always
        /// yield the domain the workstation is joined to, and the other farm answers
        /// with 401.
        /// </summary>
        public static string GetDomainFromUrl(string siteUrl)
        {
            if (string.IsNullOrWhiteSpace(siteUrl)) return string.Empty;

            try
            {
                var parts = new Uri(siteUrl).Host.Split('.');
                return parts.Length > 1 ? parts[1].ToLowerInvariant() : string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }


		/// <summary>
		/// Reads the stored user name (Param1) for the domain that owns the given
		/// site — no DPAPI decryption of the password — and returns it prefixed
		/// with the domain (e.g. "ADA\ashersa"), so the same-looking account name
		/// on two different farms isn't ambiguous in the UI. Returns null when no
		/// profile exists or the domain cannot be resolved from the URL.
		/// </summary>
		public static string? GetStoredUsername(string siteUrl)
		{
			string domain = GetDomainFromUrl(siteUrl);
			if (string.IsNullOrEmpty(domain)) return null;

			using (var key = Registry.CurrentUser.OpenSubKey(ProfilePath(domain)))
			{
				string? userName = key?.GetValue("Param1")?.ToString();
				return string.IsNullOrEmpty(userName) ? null : $@"{domain.ToUpperInvariant()}\{userName}";
			}
		}
        /// <summary>
        /// Reads the stored credentials for the domain that owns the given site.
        /// Returns null when no profile exists — the caller is expected to prompt.
        /// </summary>
        public static NetworkCredential? GetCredentials(string siteUrl)
        {
            string domain = GetDomainFromUrl(siteUrl);
            if (string.IsNullOrEmpty(domain))
            {
                _log.Caller().Warning("GetCredentials — cannot resolve a domain from '{Url}'", siteUrl);
                return null;
            }

            using (var key = Registry.CurrentUser.OpenSubKey(ProfilePath(domain)))
            {
                if (key == null)
                {
                    _log.Debug("GetCredentials — no profile for domain '{Domain}'", domain);
                    return null;
                }

                var userName     = key.GetValue("Param1")?.ToString();
                var encryptedHex = key.GetValue("Param")?.ToString();

                if (string.IsNullOrEmpty(userName) || string.IsNullOrEmpty(encryptedHex))
                {
                    _log.Caller().Warning("GetCredentials — profile '{Domain}' is incomplete", domain);
                    return null;
                }

                try
                {
                    _log.Debug("GetCredentials — '{Domain}\\{User}' → {Url}", domain, userName, siteUrl);
                    return new NetworkCredential(userName, DecryptFromPowerShell(encryptedHex), domain);
                }
                catch (Exception ex)
                {
                    // The DPAPI blob belongs to another Windows account or machine.
                    _log.Caller().Warning(ex, "GetCredentials — stored password for '{Domain}' cannot be decrypted", domain);
                    return null;
                }
            }
        }

        /// <summary>True when a usable profile exists for the domain of the given site.</summary>
        public static bool HasCredentials(string siteUrl) => GetCredentials(siteUrl) != null;

        /// <summary>
        /// Stores credentials for the domain that owns the given site. The password is
        /// protected with DPAPI in the same format PowerShell's ConvertFrom-SecureString
        /// produces, so profiles written by either tool are readable by both.
        /// </summary>
        public static void SaveCredentials(string siteUrl, string userName, string password)
        {
            string domain = GetDomainFromUrl(siteUrl);
            if (string.IsNullOrEmpty(domain))
                throw new InvalidOperationException(
                    $"Cannot determine the AD domain from the site URL:\n{siteUrl}");

            byte[] data      = Encoding.Unicode.GetBytes(password);
            byte[] encrypted = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
            string hex       = BitConverter.ToString(encrypted).Replace("-", "");

            using (var key = Registry.CurrentUser.CreateSubKey(ProfilePath(domain)))
            {
                key.SetValue("Param1", userName);
                key.SetValue("Param",  hex);
            }

            _log.Information("SaveCredentials — profile stored for domain '{Domain}', user '{User}'",
                domain, userName);
        }

    }
}