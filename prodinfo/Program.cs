// Based on http://stackoverflow.com/a/17936065/1761490

using System;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace prodinfo {
    class Program {

        private const string UpgradeCodeRegistryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes";
        private const string UninstallCodeRegistryKey = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall";
        private const string Uninstall32CodeRegistryKey = @"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall";

        private static readonly int[] GuidRegistryFormatPattern = { 8, 4, 4, 2, 2, 2, 2, 2, 2, 2, 2 };

        public static void DumpAllUpgradeCodes() {
            // Open the upgrade code registry key
            var upgradeCodeRegistryRoot = Registry.LocalMachine.OpenSubKey(UpgradeCodeRegistryKey);

            if (upgradeCodeRegistryRoot == null) {
                return;
            }
            // Iterate over each sub-key
            foreach (var subKeyName in upgradeCodeRegistryRoot.GetSubKeyNames()) {
                var subkey = upgradeCodeRegistryRoot.OpenSubKey(subKeyName);
                if (subkey == null) {
                    continue;
                }
                var formattedUpgradeCode = subkey.Name.Split('\\').LastOrDefault();
                Console.WriteLine(ConvertFromRegistryFormat(formattedUpgradeCode));

                // Check for a value containing the product code
                foreach (var productkey in subkey.GetValueNames()) {
                    // Extract the name of the subkey from the qualified name
                    var guid = ConvertFromRegistryFormat(productkey);
                    Console.WriteLine("    " + guid + " " + ProductName(guid));
                }
            }
        }

        public static Guid[] GetProductCode(Guid upgradeCode) {
            // Convert the upgrade code to the format found in the registry
            var upgradeCodeSearchString = ConvertToRegistryFormat(upgradeCode);
            // try to open the registry key
            var upgradeCodeRegistryKey = Registry.LocalMachine.OpenSubKey(UpgradeCodeRegistryKey + @"\" + upgradeCodeSearchString);
            return upgradeCodeRegistryKey == null ? null : upgradeCodeRegistryKey.GetValueNames().Select(ConvertFromRegistryFormat).ToArray<Guid>();
        }

        public static string ProductName(Guid productCode) {
            var keyname  = UninstallCodeRegistryKey + @"\{" + productCode + "}";
            var uninstallProductKey = Registry.LocalMachine.OpenSubKey(keyname);
            if (uninstallProductKey == null) {
                // try 32 products
                keyname = Uninstall32CodeRegistryKey + @"\{" + productCode + "}" ;
                uninstallProductKey = Registry.LocalMachine.OpenSubKey(keyname);
                if (uninstallProductKey == null) {
                    return "[undefined]";
                }
            }
            return (string)uninstallProductKey.GetValue("DisplayName");
        }

        public static Guid? GetUpgradeCode(Guid productCode) {
            // Convert the product code to the format found in the registry
            var productCodeSearchString = ConvertToRegistryFormat(productCode);

            // Open the upgrade code registry key
            var upgradeCodeRegistryRoot = Registry.LocalMachine.OpenSubKey(UpgradeCodeRegistryKey);

            if (upgradeCodeRegistryRoot == null) {
                return null;
            }
            // Iterate over each sub-key
            foreach (var subKeyName in upgradeCodeRegistryRoot.GetSubKeyNames()) {
                var subkey = upgradeCodeRegistryRoot.OpenSubKey(subKeyName);

                if (subkey == null) {
                    continue;
                }
                // Check for a value containing the product code
                if (subkey.GetValueNames().Any(s => s.IndexOf(productCodeSearchString, StringComparison.OrdinalIgnoreCase) >= 0)) {
                    // Extract the name of the subkey from the qualified name
                    var formattedUpgradeCode = subkey.Name.Split('\\').LastOrDefault();

                    // Convert it back to a Guid
                    return ConvertFromRegistryFormat(formattedUpgradeCode);
                }
            }
            return null;
        }

        private static string ConvertToRegistryFormat(Guid productCode) {
            return Reverse(productCode, GuidRegistryFormatPattern);
        }

        private static Guid ConvertFromRegistryFormat(string upgradeCode) {
            if (upgradeCode == null || upgradeCode.Length != 32) {
                throw new FormatException("Product code was in an invalid format");
            }
            upgradeCode = Reverse(upgradeCode, GuidRegistryFormatPattern);
            return Guid.Parse(upgradeCode);
        }

        private static string Reverse(object value, params int[] pattern) {
            // Strip the hyphens
            var inputString = value.ToString().Replace("-", "");
            var returnString = new StringBuilder();
            var index = 0;

            // Iterate over the reversal pattern
            foreach (var length in pattern) {
                // Reverse the sub-string and append it
                returnString.Append(inputString.Substring(index, length).Reverse().ToArray());

                // Increment our posistion in the string
                index += length;
            }
            return returnString.ToString();
        }

        static void PrintHelp() {
            Console.WriteLine("Usage: prodinfo [upgrade-id]");
            Console.WriteLine("  procinfo - dump product guid for a given upgrade-id");
            Console.WriteLine("  without arguments list all updrade-ids and corresponding products");
        }

        static int Main(string[] args) {

            if (args.Length == 0) {
                DumpAllUpgradeCodes();
                return 0;
            }
            if (args.Length != 1) {
                Console.WriteLine("too many arguments");
                return 2;
            }
            if (args[0].ToLower() == "-h") {
                PrintHelp();
                return 0;
            }
            Guid upgradeGuid;
            if ( ! Guid.TryParse(args[0], out upgradeGuid)) {
                Console.WriteLine("Error: failed to convert argument to guid");
                Console.WriteLine();
                PrintHelp();
                return 1;
            }
            foreach (var guid in GetProductCode(Guid.Parse(args[0]))) {
                Console.WriteLine("{0} {1}", guid, ProductName(guid));
            }
            return 0;
        }
    }
}
