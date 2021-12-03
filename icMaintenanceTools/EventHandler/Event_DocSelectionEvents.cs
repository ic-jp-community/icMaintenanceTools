using System;
using System.Collections.Generic;
using System.Text;

using interop.ICApiIronCAD;
using System.Windows.Forms;

namespace ICApiAddin.icMaintenanceTools
{
    class Event_DocSelectionEvents
    {
        private CZSelectorEvents m_selectionEvents;
        internal Event_DocSelectionEvents(CZSelectorEvents selectionEvents)
        {
            if (selectionEvents != null)
            {
                m_selectionEvents = selectionEvents;
                m_selectionEvents.SetSelectionFilterChoices((int)eZSelectionType.Z_SEL_ANY, eZSelectionType.Z_SEL_ANY);
                m_selectionEvents.Selected += new _IZSelectEvents_SelectedEventHandler(m_selectionEvents_Selected);
                m_selectionEvents.DeSelected += new _IZSelectEvents_DeSelectedEventHandler(m_selectionEvents_DeSelected);
                m_selectionEvents.DeSelectedAll += new _IZSelectEvents_DeSelectedAllEventHandler(m_selectionEvents_DeSelectedAll);
            }
        }

        void m_selectionEvents_DeSelectedAll()
        {
            MessageBox.Show("DeSelectedAllイベント");
        }

        void m_selectionEvents_DeSelected(IZElement piElement, ZMathPoint piModelCoord, int lXWindowPixel, int lYWindowPixel, int lEFlags, eZEntityType eEntType, object varFaceIds)
        {
            MessageBox.Show("DeSelectedイベント");
        }

        void m_selectionEvents_Selected(IZElement piElement, ZMathPoint piModelCoord, int lXWindowPixel, int lYWindowPixel, int lEFlags, eZEntityType eEntType, object varEntIds)
        {
            MessageBox.Show("Selectedイベント");
        }
    }

}
