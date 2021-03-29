using System;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using OBSWebsocketDotNet;
using System.Text.RegularExpressions;

namespace ObsPowerPointIntegration.Vsto
{
    public partial class ObsPowerPointIntegrationAddIn
    {
        private readonly OBSWebsocket _obsWebSocket = new OBSWebsocket();

        private string ObsHost { get; set; }
        private string ObsPassword { get; set; }
        private string ObsSceneRegex { get; set; }

        private void InternalStartup()
        {
            Startup += new EventHandler(ObsPowerPointIntegrationAddIn_Startup);
            Shutdown += new EventHandler(ObsPowerPointIntegrationAddIn_Shutdown);
            Application.SlideShowNextSlide += Application_SlideShowNextSlide;
        }

        private void ObsPowerPointIntegrationAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                ObsHost = RegistryHelper.GetString("OBSHost");
                ObsPassword = RegistryHelper.GetString("OBSPassword");
                ObsSceneRegex = RegistryHelper.GetString("OBSSceneRegex");

                _obsWebSocket.Connect(ObsHost, ObsPassword);
                var versionInfo = _obsWebSocket.GetVersion();
                Debug.WriteLine($"Connected to OBS {versionInfo.OBSStudioVersion}");
            }
            catch
            {
                Debug.WriteLine("Connection to OBS failed!");
            }
        }

        private void ObsPowerPointIntegrationAddIn_Shutdown(object sender, EventArgs e)
        {
            if (_obsWebSocket != null && _obsWebSocket.IsConnected)
            {
                _obsWebSocket.Disconnect();
                Debug.WriteLine($"Disconnected from OBS");
            }
        }

        private void Application_SlideShowNextSlide(SlideShowWindow window)
        {
            if (window != null)
            {
                try
                {
                    string slideName = window.View.Slide.Name;
                    string slideNotes = window.View.Slide.NotesPage.Shapes[2].TextFrame.TextRange.Text;
                    if (string.IsNullOrWhiteSpace(slideNotes))
                    {
                        Debug.WriteLine($"{slideName} DOES NOT contain any notes!");
                    }
                    else
                    {
                        Debug.WriteLine($"{slideName} contains notes: {slideNotes}!");
                    }

                    string obsScene = ExpandOBSScene(slideNotes);
                    if (string.IsNullOrWhiteSpace(slideNotes))
                    {
                        Debug.WriteLine($"No OBS Scene placeholder found for {slideName}!");
                    }
                    else
                    {
                        Debug.WriteLine($"OBS Scene {obsScene} found for {slideName}!");
                        _obsWebSocket.SetCurrentScene(obsScene);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Something ('{ex.Message}') went wrong while reacting to showing next slide :(");
                }
            }
        }

        private string ExpandOBSScene(string slideNotes)
        {
            // {OBS:([\s\S]+)}
            var regexMatch = Regex.Match(slideNotes, ObsSceneRegex);
            if (regexMatch.Success && regexMatch.Groups.Count == 2)
            {
                return regexMatch.Groups[1].Value;
            }

            return string.Empty;
        }
    }
}