#region Header
//
//Copyright(c) 2019 Timothy Logan, HKS Inc

//Permission is hereby granted, free of charge, to any person obtaining
//a copy of this software and associated documentation files (the
//"Software"), to deal in the Software without restriction, including
//without limitation the rights to use, copy, modify, merge, publish,
//distribute, sublicense, and/or sell copies of the Software, and to
//permit persons to whom the Software is furnished to do so, subject to
//the following conditions:

//The above copyright notice and this permission notice shall be
//included in all copies or substantial portions of the Software.

//THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
//EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
//MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
//NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
//LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
//OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
//WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
#endregion

using System.Collections.Generic;
using System.Linq;
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
