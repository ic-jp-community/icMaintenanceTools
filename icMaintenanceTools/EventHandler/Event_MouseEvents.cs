using System;
using System.Collections.Generic;
using System.Text;

using interop.ICApiIronCAD;
using System.Windows.Forms;


namespace ICApiAddin.icMaintenanceTools
{

    class Event_MouseEvents
    {
        internal Event_MouseEvents(ZMouseEvents zMouseEvent)
        {
            if (zMouseEvent != null)
            {
                zMouseEvent.OnLButtonDown += new _IZMouseEvents_OnLButtonDownEventHandler(m_zMouseEvents_OnLButtonDown);
                zMouseEvent.OnLButtonDblClick += new _IZMouseEvents_OnLButtonDblClickEventHandler(m_zMouseEvents_OnLButtonDblClick);
            }
        }

        void m_zMouseEvents_OnLButtonDblClick(int lClientX, int lClientY, int lKeyState)
        {
            MessageBox.Show("Left mouse button double clickイベント");
        }

        void m_zMouseEvents_OnLButtonDown(int lClientX, int lClientY, int lKeyState)
        {
            MessageBox.Show("Left mouse button downイベント");
        }
    }


}
