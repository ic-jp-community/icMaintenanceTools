using System;
using System.Collections.Generic;
using System.Text;

using interop.ICApiIronCAD;
using System.Collections;

namespace ICApiAddin.icMaintenanceTools
{
    /// <summary>
    /// Utility class conatins static utility functions
    /// </summary>
    static class Event_ICAddInUtils
    {
        internal static List<IZDoc> GetAllOpenDocs(IZBaseApp izBaseApp)
        {
            List<IZDoc> list = new List<IZDoc>();
            if (izBaseApp != null)
            {
                object vDocs = izBaseApp.GetAllOpenDocs();
                if (vDocs != null)
                {
                    //Process VARIANT of oDocs to IZoc array/list
                    IEnumerable iIZDocs = vDocs as IEnumerable;
                    if (iIZDocs != null)
                    {
                        foreach (object oDoc in iIZDocs)
                        {
                            IZDoc iZDoc = oDoc as IZDoc;
                            if (iZDoc != null)
                            {
                                list.Add(iZDoc);
                            }
                        }
                    }
                }
            }
            return list;
        }
    }
}
