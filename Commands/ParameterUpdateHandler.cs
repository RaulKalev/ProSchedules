using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ProSchedules.Commands
{
    public class ParameterUpdateHandler : IExternalEventHandler
    {
        // Tuple: ElementId, ParameterName, NewValue
        public List<Tuple<string, string, string>> Updates { get; set; } = new List<Tuple<string, string, string>>();

        public void Execute(UIApplication app)
        {
            if (Updates == null || Updates.Count == 0) return;

            Document doc = app.ActiveUIDocument.Document;

            using (Transaction t = new Transaction(doc, "Batch Parameter Update"))
            {
                t.Start();

                int successCount = 0;
                int failCount = 0;

                foreach (var update in Updates)
                {
                    try
                    {
                        string idStr = update.Item1;
                        string paramName = update.Item2;
                        string newValue = update.Item3;

                        // Parse ElementId
                        ElementId id = ElementId.InvalidElementId;
                        if (long.TryParse(idStr, out long idVal))
                            id = new ElementId(idVal);

                        if (id == ElementId.InvalidElementId) continue;

                        Element el = doc.GetElement(id);
                        if (el == null) continue;

                        Parameter param = el.LookupParameter(paramName);
                        if (param == null)
                        {
                            // Try type parameter logic if needed, but for now strict lookup
                            // Logic from RevitService fallback could be reused here or strictly separate
                            ElementId typeId = el.GetTypeId();
                            if (typeId != ElementId.InvalidElementId)
                            {
                                Element typeElem = doc.GetElement(typeId);
                                if (typeElem != null)
                                    param = typeElem.LookupParameter(paramName);
                            }
                        }

                        if (param != null && !param.IsReadOnly)
                        {
                            if (param.StorageType == StorageType.String)
                            {
                                param.Set(newValue);
                                successCount++;
                            }
                            else if (param.StorageType == StorageType.Double || param.StorageType == StorageType.Integer)
                            {
                                // Simple ValueString attempt (Set doesn't take ValueString usually, 
                                // typically SetValueString is for Dimension types, 
                                // generic Set takes double/int. 
                                // BUT parsing string to double/int is risky without units.
                                // For now, let's try SetValueString api if available or fallbacks.
                                if (param.SetValueString(newValue))
                                {
                                    successCount++;
                                }
                                else
                                {
                                    // Try raw parse?
                                    failCount++;
                                }
                            }
                        }
                    }
                    catch
                    {
                        failCount++;
                    }
                }

                t.Commit();
            }

            Updates.Clear();
        }

        public string GetName()
        {
            return "ProSchedules Parameter Update";
        }
    }
}
