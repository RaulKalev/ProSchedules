using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ProSchedules.Models;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ProSchedules.Services
{
    public class RevitService
    {
        private Document _doc;

        public RevitService(Document doc)
        {
            _doc = doc;
        }

        public List<SheetItem> GetSheets()
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(s => !s.IsTemplate)
                .Select(s => new SheetItem(s))
                .OrderBy(s => s.SheetNumber)
                .ToList();
        }

        public List<ViewSchedule> GetSchedules()
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(s => !s.IsTemplate && !s.IsInternalKeynoteSchedule && !s.IsTitleblockRevisionSchedule)
                .OrderBy(s => s.Name)
                .ToList();
        }

        public ScheduleData GetScheduleData(ViewSchedule schedule)
        {
            var data = new Models.ScheduleData();
            data.ScheduleId = schedule.Id;
            if (schedule == null) return data;

            ScheduleDefinition def = schedule.Definition;
            ElementId categoryId = def.CategoryId;

            var fields = new List<ScheduleField>();
            var fieldIds = def.GetFieldOrder();

            foreach (var id in fieldIds)
            {
                ScheduleField field = def.GetField(id);
                if (!field.IsHidden)
                {
                    fields.Add(field);
                    data.Columns.Add(field.GetName());
                }
            }

            IList<Element> elements = new List<Element>();
            if (categoryId != ElementId.InvalidElementId)
            {
                elements = new FilteredElementCollector(_doc)
                    .OfCategoryId(categoryId)
                    .WhereElementIsNotElementType()
                    .ToElements();
            }

            if (!data.Columns.Contains("ElementId"))
            {
                data.Columns.Insert(0, "ElementId");
            }
            if (!data.Columns.Contains("TypeName"))
            {
                data.Columns.Insert(1, "TypeName");
            }

            foreach (Element el in elements)
            {
                var rowData = new List<string>();
#if NET8_0_OR_GREATER
                rowData.Add(el.Id.Value.ToString());
#else
                rowData.Add(el.Id.IntegerValue.ToString());
#endif
                
                ElementId typeId = el.GetTypeId();
                string typeName = "";
                if (typeId != ElementId.InvalidElementId)
                {
                    Element typeElem = _doc.GetElement(typeId);
                    if (typeElem != null) typeName = typeElem.Name;
                }
                rowData.Add(typeName);

                foreach (ScheduleField field in fields)
                {
                    string val = "";
                    bool isType = false;
                    
                    if (field.ParameterId == ElementId.InvalidElementId)
                    {
                        val = "";
                    }
                    else
                    {
                        var result = GetParameterValue(el, field.ParameterId);
                        val = result.Item1;
                        isType = result.Item2;
                    }
                    
                    if (val == null) val = "";
                    rowData.Add(val);

                    string colName = field.GetName();
                    if (!data.IsTypeParameter.ContainsKey(colName))
                    {
                        data.IsTypeParameter[colName] = isType;
                    }
                    if (isType) data.IsTypeParameter[colName] = true;
                }
                data.Rows.Add(rowData);
            }

            return data;
        }

        private (string, bool) GetParameterValue(Element el, ElementId parameterId)
        {
            Parameter p = null;
            bool isType = false;
            
#if NET8_0_OR_GREATER
            long idValue = parameterId.Value;
#else
            int idValue = parameterId.IntegerValue;
#endif

            if (idValue < 0)
            {
                p = el.get_Parameter((BuiltInParameter)idValue);
            }
            else
            {
                try
                {
                    var paramElem = _doc.GetElement(parameterId);
                    if (paramElem != null)
                    {
                        p = el.LookupParameter(paramElem.Name);
                    }
                }
                catch { }
            }

            if (p == null)
            {
                ElementId typeId = el.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    Element typeElem = _doc.GetElement(typeId);
                    if (typeElem != null)
                    {
                         if (idValue < 0)
                            p = typeElem.get_Parameter((BuiltInParameter)idValue);
                         else
                         {
                             var paramElem = _doc.GetElement(parameterId);
                             if (paramElem != null) p = typeElem.LookupParameter(paramElem.Name);
                         }
                         
                         if (p != null) isType = true;
                    }
                }
            }

            if (p != null)
            {
                string val;
                if (p.StorageType == StorageType.String)
                    val = p.AsString();
                else
                    val = p.AsValueString();
                
                return (val, isType);
            }

            return ("", false);
        }
    }
}
