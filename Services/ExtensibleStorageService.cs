using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using System;
using System.Linq;

namespace ProSchedules.Services
{
    /// <summary>
    /// Reads and writes per-user sort/filter settings to Revit Extensible Storage.
    /// Each user owns their own DataStorage element, preventing worksharing conflicts.
    /// </summary>
    internal static class ExtensibleStorageService
    {
        // Stable GUID — never change once deployed
        private static readonly Guid SchemaGuid = new Guid("3F8A2C1D-B47E-4F90-A523-6D1E87C4F205");
        private const string SchemaName       = "ProSchedulesUserSettings";
        private const string FieldSort          = "SortSettingsJson";
        private const string FieldFilter        = "FilterSettingsJson";
        private const string DataStoragePrefix  = "ProSchedules_Settings_";

        // ------------------------------------------------------------------ //
        //  Schema                                                              //
        // ------------------------------------------------------------------ //

        private static Schema GetOrCreateSchema()
        {
            Schema schema = Schema.Lookup(SchemaGuid);
            if (schema != null) return schema;

            var builder = new SchemaBuilder(SchemaGuid);
            builder.SetSchemaName(SchemaName);
            builder.SetReadAccessLevel(AccessLevel.Public);
            builder.SetWriteAccessLevel(AccessLevel.Public);
            builder.AddSimpleField(FieldSort,         typeof(string));
            builder.AddSimpleField(FieldFilter,       typeof(string));
            return builder.Finish();
        }

        // ------------------------------------------------------------------ //
        //  DataStorage lookup                                                  //
        // ------------------------------------------------------------------ //

        private static string SanitizeName(string username)
        {
            // Keep only safe characters for element names
            var safe = new System.Text.StringBuilder();
            foreach (char c in username)
                safe.Append(char.IsLetterOrDigit(c) || c == '_' || c == '-' ? c : '_');
            return safe.ToString();
        }

        private static DataStorage FindDataStorage(Document doc, string username)
        {
            string name = DataStoragePrefix + SanitizeName(username);
            return new FilteredElementCollector(doc)
                .OfClass(typeof(DataStorage))
                .Cast<DataStorage>()
                .FirstOrDefault(ds => ds.Name == name);
        }

        /// <summary>
        /// Must be called inside an open Transaction.
        /// </summary>
        private static DataStorage GetOrCreateDataStorage(Document doc, string username)
        {
            var ds = FindDataStorage(doc, username);
            if (ds != null) return ds;

            ds = DataStorage.Create(doc);
            ds.Name = DataStoragePrefix + SanitizeName(username);
            return ds;
        }

        // ------------------------------------------------------------------ //
        //  Public API                                                          //
        // ------------------------------------------------------------------ //

        /// <summary>
        /// Loads sort and filter JSON strings for the given user. Safe to call
        /// outside a transaction (reads don't require one).
        /// Returns (null, null) if no settings have been saved yet.
        /// </summary>
        public static (string sortJson, string filterJson) LoadSettings(Document doc, string username)
        {
            try
            {
                var ds = FindDataStorage(doc, username);
                if (ds == null) return (null, null);

                // Schema must be registered each Revit session — GetOrCreateSchema handles both
                // the first-time-ever case and the "new session, schema not in memory" case.
                Schema schema = GetOrCreateSchema();

                Entity entity = ds.GetEntity(schema);
                if (!entity.IsValid()) return (null, null);

                string sortJson   = entity.Get<string>(schema.GetField(FieldSort));
                string filterJson = entity.Get<string>(schema.GetField(FieldFilter));
                return (sortJson, filterJson);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ExtensibleStorageService] LoadSettings error: {ex.Message}");
                return (null, null);
            }
        }

        /// <summary>
        /// Persists sort and filter JSON strings for the given user.
        /// MUST be called inside an open Transaction.
        /// </summary>
        public static void SaveSettings(Document doc, string username, string sortJson, string filterJson)
        {
            Schema schema = GetOrCreateSchema();
            DataStorage ds = GetOrCreateDataStorage(doc, username);

            Entity entity = new Entity(schema);
            entity.Set(schema.GetField(FieldSort),   sortJson   ?? string.Empty);
            entity.Set(schema.GetField(FieldFilter), filterJson ?? string.Empty);
            ds.SetEntity(entity);
        }
    }
}
