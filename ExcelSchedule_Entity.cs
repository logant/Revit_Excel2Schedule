using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using RevitCommon;
using RevitCommon.Attributes;

namespace LINE.Revit
{
    [Schema("91c053bd-edeb-4feb-abac-ef862c311e9d", "ExcelSchedules", Documentation = "Data to manage Excel links", 
        ReadAccessLevel = AccessLevel.Public, VendorId = "HKSL", WriteAccessLevel = AccessLevel.Vendor)]
    public class ExcelScheduleEntity : IRevitEntity
    {
        [Field(Documentation = "Schedule Element ID")]
        public List<ElementId> ScheduleId { get; set; }

        [Field(Documentation = "Excel Schedule Path")]
        public List<string> ExcelFilePath { get; set; }

        [Field(Documentation = "Excel Schedule Path Type")]
        public List<int> PathType { get; set; }

        [Field(Documentation = "Excel Worksheet Name")]
        public List<string> WorksheetName { get; set; }

        [Field(Documentation = "Excel Last Modify Date")]
        public List<string> DateTime { get; set; }

        public static bool VerifySchema()
        {
            if (!(typeof(ExcelScheduleEntity).GetCustomAttributes(typeof(SchemaAttribute), true).FirstOrDefault() is SchemaAttribute schemaAttr))
                return false;

            Schema schemaCheck = Schema.Lookup(schemaAttr.GUID);
            return null != schemaCheck && schemaCheck.IsValidObject;
        }
    }
}
