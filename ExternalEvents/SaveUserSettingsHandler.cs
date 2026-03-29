using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ProSchedules.Services;
using System;

namespace ProSchedules.ExternalEvents
{
    /// <summary>
    /// Writes sort and filter settings to Revit Extensible Storage.
    /// Must be raised via ExternalEvent so it runs on Revit's main thread with
    /// access to a valid API context and transaction.
    /// </summary>
    internal class SaveUserSettingsHandler : IExternalEventHandler
    {
        public string SortSettingsJson   { get; set; }
        public string FilterSettingsJson { get; set; }

        public void Execute(UIApplication app)
        {
            try
            {
                Document doc      = app.ActiveUIDocument?.Document;
                string   username = app.Application.Username;

                if (doc == null || doc.IsReadOnly) return;

                using (var tx = new Transaction(doc, "Save ProSchedules Settings"))
                {
                    tx.Start();
                    ExtensibleStorageService.SaveSettings(doc, username, SortSettingsJson, FilterSettingsJson);
                    tx.Commit();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SaveUserSettingsHandler] Execute error: {ex.Message}");
            }
        }

        public string GetName() => "ProSchedules Save User Settings";
    }
}
