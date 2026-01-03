using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;

namespace ProSchedules.ExternalEvents
{
    public class ParameterValueUpdateHandler : IExternalEventHandler
    {
        public string ElementIdStr { get; set; }
        public string ParameterIdStr { get; set; }
        public string NewValue { get; set; }
        public bool Success { get; private set; }
        public string ErrorMessage { get; private set; }

        public void Execute(UIApplication app)
        {
            Success = false;
            ErrorMessage = "";

            try
            {
                Document doc = app.ActiveUIDocument.Document;

                if (!long.TryParse(ElementIdStr, out long elemIdValue))
                {
                    ErrorMessage = "Invalid element ID";
                    return;
                }

                if (!long.TryParse(ParameterIdStr, out long paramIdValue))
                {
                    ErrorMessage = "Invalid parameter ID";
                    return;
                }

                ElementId elementId = new ElementId(elemIdValue);
                ElementId parameterId = new ElementId(paramIdValue);

                Element element = doc.GetElement(elementId);
                if (element == null)
                {
                    ErrorMessage = "Element not found";
                    return;
                }

                using (Transaction trans = new Transaction(doc, "Update Parameter Value"))
                {
                    trans.Start();

                    bool updated = SetParameterValue(doc, element, parameterId, NewValue);

                    if (updated)
                    {
                        trans.Commit();
                        Success = true;
                    }
                    else
                    {
                        trans.RollBack();
                        ErrorMessage = "Failed to set parameter value";
                    }
                }
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
            }
        }

        private bool SetParameterValue(Document doc, Element element, ElementId parameterId, string value)
        {
            Parameter p = null;
            long idValue = parameterId.Value;

            // Try to get parameter from instance
            if (idValue < 0)
            {
                p = element.get_Parameter((BuiltInParameter)(int)idValue);
            }
            else
            {
                try
                {
                    var paramElem = doc.GetElement(parameterId);
                    if (paramElem != null)
                    {
                        p = element.LookupParameter(paramElem.Name);
                    }
                }
                catch { }
            }

            // Fallback: iterate through instance parameters
            if (p == null)
            {
                foreach (Parameter param in element.Parameters)
                {
                    if (param.Id.Value == parameterId.Value)
                    {
                        p = param;
                        break;
                    }
                }
            }

            // If not found on instance, try type
            if (p == null)
            {
                ElementId typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    Element typeElem = doc.GetElement(typeId);
                    if (typeElem != null)
                    {
                        if (idValue < 0)
                        {
                            p = typeElem.get_Parameter((BuiltInParameter)(int)idValue);
                        }
                        else
                        {
                            var paramElem = doc.GetElement(parameterId);
                            if (paramElem != null)
                            {
                                p = typeElem.LookupParameter(paramElem.Name);
                            }
                        }

                        if (p == null)
                        {
                            foreach (Parameter param in typeElem.Parameters)
                            {
                                if (param.Id.Value == parameterId.Value)
                                {
                                    p = param;
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            if (p == null || p.IsReadOnly)
                return false;

            try
            {
                switch (p.StorageType)
                {
                    case StorageType.String:
                        p.Set(value);
                        return true;

                    case StorageType.Integer:
                        if (int.TryParse(value, out int intVal))
                        {
                            p.Set(intVal);
                            return true;
                        }
                        break;

                    case StorageType.Double:
                        if (double.TryParse(value, out double dblVal))
                        {
                            p.Set(dblVal);
                            return true;
                        }
                        break;

                    case StorageType.ElementId:
                        // For now, skip ElementId type parameters
                        return false;
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        public string GetName()
        {
            return "ParameterValueUpdateHandler";
        }
    }
}
