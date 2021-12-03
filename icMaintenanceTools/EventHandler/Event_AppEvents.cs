
using System.Collections.Generic;
using interop.ICApiIronCAD;
using System.Windows.Forms;

namespace ICApiAddin.icMaintenanceTools
{
    class AppEvents
    {
        private ZIronCADApp m_zIronCADApp;

        internal AppEvents()
        {
        }

        internal void RegisterForAppEvents(IZBaseApp izBaseApp)
        {
            m_zIronCADApp = (ZIronCADApp)izBaseApp;
            if (m_zIronCADApp != null)
            {
                m_zIronCADApp.ActiveDocChanged += new _IZApplicationEvents_ActiveDocChangedEventHandler(m_zIronCADApp_ActiveDocChanged);
                m_zIronCADApp.DocumentOpened += new _IZApplicationEvents_DocumentOpenedEventHandler(m_zIronCADApp_DocumentOpened);
                m_zIronCADApp.DocumentCreated += new _IZApplicationEvents_DocumentCreatedEventHandler(m_zIronCADApp_DocumentCreated);
                m_zIronCADApp.DocumentPreClose += new _IZApplicationEvents_DocumentPreCloseEventHandler(m_zIronCADApp_DocumentPreClose);
            }

            List<IZDoc> iZDocList = Event_ICAddInUtils.GetAllOpenDocs(izBaseApp);
            foreach (IZDoc iZDoc in iZDocList)
            {
                if (iZDoc != null)
                {
                    RegisterForDocEvents(iZDoc);
                }
            }
        }

        internal void UnRegisterForAppEvents()
        {
            if (m_zIronCADApp != null)
            {
                m_zIronCADApp.ActiveDocChanged -= m_zIronCADApp_ActiveDocChanged;
                m_zIronCADApp.DocumentCreated -= m_zIronCADApp_DocumentCreated;
                m_zIronCADApp.DocumentOpened -= m_zIronCADApp_DocumentOpened;
                m_zIronCADApp.DocumentPreClose -= m_zIronCADApp_DocumentPreClose;

                IZBaseApp izBaseApp = m_zIronCADApp as IZBaseApp;
                if (izBaseApp != null)
                {
                    List<IZDoc> iZDocList = Event_ICAddInUtils.GetAllOpenDocs(izBaseApp);
                    foreach (IZDoc iZDoc in iZDocList)
                    {
                        if (iZDoc != null)
                        {
                            UnRegisterForDocEvents(iZDoc);
                        }
                    }
                }
            }
        }

        void m_zIronCADApp_DocumentPreClose(IZDoc pNewDoc)
        {
            MessageBox.Show("ファイルクローズ前(PreClose)イベント");
            if (pNewDoc != null)
            {
                UnRegisterForDocEvents(pNewDoc);
            }
        }

        void m_zIronCADApp_DocumentCreated(IZDoc piNewDoc)
        {
            MessageBox.Show("ファイル作成(Created)イベント");
            if (piNewDoc != null)
            {
                RegisterForDocEvents(piNewDoc);
            }
        }

        void m_zIronCADApp_DocumentOpened(IZDoc piNewDoc)
        {
            MessageBox.Show("ファイルを開く(Opened)イベント");
            if (piNewDoc != null)
            {
                RegisterForDocEvents(piNewDoc);
            }
        }

        void m_zIronCADApp_ActiveDocChanged(IZDoc piNewDoc)
        {
            MessageBox.Show("アクティブファイルの変更(ActiveDocChanged)イベント");
        }

        internal void RegisterForDocEvents(IZDoc iZDoc)
        {
            if (iZDoc != null)
            {
                IZSceneDoc iZsceneDoc = iZDoc as IZSceneDoc;
                IZSelectionMgr iZselectionMgr = iZsceneDoc.SelectionMgr;
                CZInteractor cZinteractor = iZselectionMgr.CreateInteractor();

                //Doc events
                Event_DocEvents docEvents = new Event_DocEvents();
                docEvents.RegisterEvents(iZDoc);

                //Handle Events
                ZHandleEvents zHandleEvents = (ZHandleEvents)iZDoc.HandleEvents;
                Event_DocHandleEvents docHandleEvents = new Event_DocHandleEvents();
                docHandleEvents.RegisterEvents(zHandleEvents);

                //Selection Events
                CZSelectorEvents cZselectionEvents = cZinteractor.SelectEvents;
                Event_DocSelectionEvents selectionEvent = new Event_DocSelectionEvents(cZselectionEvents);

                //Mouse Events
                ZMouseEvents zMouseEvents = cZinteractor.MouseEvents;
                Event_MouseEvents mouseEvents = new Event_MouseEvents(zMouseEvents);

                //Start
                cZinteractor.Start();
            }
        }

        internal void UnRegisterForDocEvents(IZDoc iZDoc)
        {
            if (iZDoc != null)
            {
                Event_DocEvents docEvents = new Event_DocEvents();
                docEvents.UnRegisterEvents(iZDoc);
                ZHandleEvents zHandleEvents = (ZHandleEvents)iZDoc.HandleEvents;
                Event_DocHandleEvents docHandleEvents = new Event_DocHandleEvents();
                docHandleEvents.UnRegisterEvents(zHandleEvents);

                IZSceneDoc iZsceneDoc = iZDoc as IZSceneDoc;
                if (iZsceneDoc != null)
                {
                    try
                    {
                        IZSelectionMgr iZselectionMgr = iZsceneDoc.SelectionMgr;
                        CZInteractor cZinteractor = iZselectionMgr.CreateInteractor();
                        cZinteractor.Stop();
                        return;
                    }
                    catch
                    {
                        //Do nothing
                    }
                }
            }
        }
    }
}
