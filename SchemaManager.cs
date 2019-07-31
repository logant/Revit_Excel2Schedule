using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using RevitCommon.Attributes;

namespace LINE.Revit
{
    public static class SchemaManager
    {
        public static DataStorage GetDataStorage(Document doc)
        {
            DataStorage ds = null;

            // See if a datastorage object already exists for this.
            ICollection<ElementId> dsCollector = new FilteredElementCollector(doc).OfClass(typeof(DataStorage)).ToElementIds();
            foreach (ElementId eid in dsCollector)
            {
                DataStorage dStor = doc.GetElement(eid) as DataStorage;
                if (dStor.Name == Properties.Settings.Default.DataStorageName)
                {
                    ds = dStor;
                    break;
                }
            }

            return ds;
        }
    }
}
