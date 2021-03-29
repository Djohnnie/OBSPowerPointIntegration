using Microsoft.Win32;

namespace ObsPowerPointIntegration.Vsto
{
    public static class RegistryHelper
    {
        private static string _keyName = "HKEY_CURRENT_USER\\Software\\Djohnnie\\OBSPowerPointIntegrationAddIn";

        public static string GetString(string key)
        {
            return (string)Registry.GetValue(_keyName, key, string.Empty);
        }
    }
}