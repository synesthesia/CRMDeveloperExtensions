using System;
using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell.Interop;

namespace ReportDeployer
{
    public sealed class VsHierarchyEvents : IVsHierarchyEvents
    {
        private readonly IVsHierarchy _hierarchy;
        private readonly ReportList _reportList;

        public VsHierarchyEvents(IVsHierarchy hierarchy, ReportList reportlist)
        {
            _hierarchy = hierarchy;
            _reportList = reportlist;
        }

        int IVsHierarchyEvents.OnInvalidateIcon(IntPtr hicon)
        {
            return VSConstants.S_OK;
        }

        int IVsHierarchyEvents.OnInvalidateItems(uint itemidParent)
        {
            return VSConstants.S_OK;
        }

        int IVsHierarchyEvents.OnItemAdded(uint itemidParent, uint itemidSiblingPrev, uint itemidAdded)
        {
            object itemExtObject;
            if (_hierarchy.GetProperty(itemidAdded, (int)__VSHPROPID.VSHPROPID_ExtObject, out itemExtObject) == VSConstants.S_OK)
            {
                var projectItem = itemExtObject as ProjectItem;
                if (projectItem != null)
                    _reportList.ProjectItemAdded(projectItem, itemidAdded);
            }
            return VSConstants.S_OK;
        }

        int IVsHierarchyEvents.OnItemDeleted(uint itemid)
        {
            object itemExtObject;
            if (_hierarchy.GetProperty(itemid, (int)__VSHPROPID.VSHPROPID_ExtObject, out itemExtObject) == VSConstants.S_OK)
            {
                var projectItem = itemExtObject as ProjectItem;
                if (projectItem != null)
                    _reportList.ProjectItemRemoved(projectItem, itemid);
            }
            return VSConstants.S_OK;
        }

        int IVsHierarchyEvents.OnItemsAppended(uint itemidParent)
        {
            return VSConstants.S_OK;
        }

        int IVsHierarchyEvents.OnPropertyChanged(uint itemid, int propid, uint flags)
        {
            //Reporting projects don't raise the normal ProjectItem Renamed event 
            //this event is raised 4 times when a report file is renamed
            //VSHPROPID_Caption happens to be the first so in order to limit the
            //item rename code from firing multiple times I'm limiting it to this property
            if (propid != (int)__VSHPROPID.VSHPROPID_Caption) return VSConstants.S_OK;

            object objProj;
            _hierarchy.GetProperty(itemid, (int)__VSHPROPID.VSHPROPID_ExtObject, out objProj);

            var projectItem = objProj as ProjectItem;

            if (projectItem != null)
                _reportList.ProjectItemRenamed(projectItem);

            return VSConstants.S_OK;
        }
    }
}
