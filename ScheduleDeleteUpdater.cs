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


using System;
using System.Collections.Generic;
using Autodesk.Revit.DB;

namespace LINE.Revit
{
    public class ScheduleDeleteUpdater : IUpdater
    {
        private readonly Guid schemaGUID = new Guid("91c053bd-edeb-4feb-abac-ef862c311e9d");
        static AddInId m_appid;
        static UpdaterId m_updaterId;

        public ScheduleDeleteUpdater(AddInId id)
        {
            m_appid = id;
            m_updaterId = new UpdaterId(m_appid, new Guid("4ba514f4-8b95-48b7-b61a-9a5552ca7b94"));
        }

        public void Execute(UpdaterData data)
        {
            Document doc = data.GetDocument();

            // Get the schema
            Autodesk.Revit.DB.ExtensibleStorage.Schema schema = Autodesk.Revit.DB.ExtensibleStorage.Schema.Lookup(schemaGUID);
            Autodesk.Revit.DB.ExtensibleStorage.Entity entity = null;
            try
            {
                entity = doc.ProjectInformation.GetEntity(schema);
            }
            catch { }

            if (entity != null)
            {
                IList<ElementId> elementIds = entity.Get<IList<ElementId>>("ScheduleId");
                if (elementIds.Count > 0)
                {

                    IList<string> paths = entity.Get<IList<string>>("ExcelFilePath");
                    IList<string> worksheets = entity.Get<IList<string>>("WorksheetName");
                    IList<string> dateTimes = entity.Get<IList<string>>("DateTime");
                    IList<int> pathTypes = entity.Get<IList<int>>("PathType");
                    foreach (ElementId deletedId in data.GetDeletedElementIds())
                    {
                        try
                        {
                            // Check if it's in the linked schedules
                            int index = -1;
                            for (int i = 0; i < elementIds.Count; i++)
                            {
                                ElementId id = elementIds[i];
                                if (id.IntegerValue == deletedId.IntegerValue || id.IntegerValue == -1)
                                {
                                    index = i;
                                }
                            }
                            if (index >= 0)
                            {
                                elementIds.RemoveAt(index);
                                worksheets.RemoveAt(index);
                                dateTimes.RemoveAt(index);
                                paths.RemoveAt(index);
                                pathTypes.RemoveAt(index);

                                // if there is still more than one element in the lists, reassign them to the entity, otherwise purge it.
                                if (elementIds.Count > 0)
                                {
                                    entity.Set("ScheduleId", elementIds);
                                    entity.Set("ExcelFilePath", paths);
                                    entity.Set("WorksheetName", worksheets);
                                    entity.Set("DateTime", dateTimes);
                                    entity.Set("PathType", pathTypes);
                                    doc.ProjectInformation.SetEntity(entity);
                                }
                                else
                                {
                                    // Delete the entity data
                                    doc.ProjectInformation.DeleteEntity(schema);
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
        }

        public string GetAdditionalInformation() { return "Check if a delete schedule was linked to Excel."; }
        public ChangePriority GetChangePriority() { return ChangePriority.Views; }
        public UpdaterId GetUpdaterId() { return m_updaterId; }
        public string GetUpdaterName() { return "ScheduleDeletedUpdater"; }
    }
}
