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

        private static string regPath = @"SOFTWARE\Microsoft\CrSiteAutomate";
        /*
		public static 	NetworkCredential GetCredentials()
		{
			
			using (var key = Registry.CurrentUser.OpenSubKey(regPath))
			{
				string userName = key?.GetValue("Param1")?.ToString() ?? "Unknown";
				string encryptedHex = key?.GetValue("Param")?.ToString() ?? "";
				return new NetworkCredential(userName, DecryptFromPowerShell(encryptedHex), "ekmd");
			}
		}
		*/
        public static NetworkCredential? GetCredentials()
        {
            

            using (var key = Registry.CurrentUser.OpenSubKey(regPath))
            {
                // Если ключа в реестре вообще нет
                if (key == null) return null;

                var userName = key.GetValue("Param1")?.ToString();
                var encryptedHex = key.GetValue("Param")?.ToString();

                // Если значения отсутствуют или пусты
                if (string.IsNullOrEmpty(userName) || string.IsNullOrEmpty(encryptedHex))
                {
                    return null;
                }

                try
                {
                    return new NetworkCredential(userName, DecryptFromPowerShell(encryptedHex), "ekmd");
                }
                catch
                {
                    // Если ошибка дешифровки (например, ключ поврежден)
                    return null;
                }
            }
        }
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
		public static void SaveCredentials(string userName, string password)
		{
			
			
			// Шифруем пароль (DPAPI)
			byte[] data = Encoding.Unicode.GetBytes(password);
			byte[] encrypted = ProtectedData.Protect(data, null, DataProtectionScope.CurrentUser);
			string hex = BitConverter.ToString(encrypted).Replace("-", "");

			using (var key = Registry.CurrentUser.CreateSubKey(regPath))
			{
				key.SetValue("Param1", userName);
				key.SetValue("Param", hex);
			}
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
            catch { /* If URL is invalid, return as-is */ }
                _log.Error("ERROR in catch block");

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
                    }).Where(attr => attr != null);

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
                        foreach (var node in xDoc.Root.Nodes())
                        {
                            node.WriteTo(xmlWriter);
                        }
                    }
                    return stringWriter.ToString();
                }
            }
            catch (Exception ex)
            {
                _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                // Если даже так не вышло, возвращаем оригинал, чтобы не падать
                System.Diagnostics.Debug.WriteLine("XML formatting error: " + ex.Message);
                return xml;
            }
        }

    }
}