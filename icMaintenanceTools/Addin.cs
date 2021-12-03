using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using interop.ICApiIronCAD;

namespace ICApiAddin.icMaintenanceTools
{
    [Guid("6AE87CEF-C966-4938-A945-40D4280F6000"), ClassInterface(ClassInterfaceType.None), ProgId("icMaintenanceTools.AddIn")]
    public class Addin : IZAddinServer
    {
         public const string ADDIN_GUID = "6AE87CEF-C966-4938-A945-40D4280F6000";

        #region [Private Members]
        private ZAddinSite m_izAddinSite;
        private ZCommandHandler m_buttonForm;
        #endregion

        //Constractor
        public Addin()
        {
#if ADDIN_INIT_DEBUG
            /* アドインロードのデバッグ用 */
            System.Threading.Thread.Sleep(120 * 1000);
#endif           
        }

#region [Public Properties]
        public IZBaseApp IronCADApp
        {
            get
            {
                if (m_izAddinSite != null)
                    return m_izAddinSite.Application;
                return null;
            }

        }

#endregion

#region [IZAddinServer Members]
        public void InitSelf(ZAddinSite piAddinSite)
        {
            if (piAddinSite != null)
            {
                m_izAddinSite = piAddinSite;
                try
                {
                    //ボタンの作成(Form)
                    stdole.IPictureDisp oImageSmall = ConvertImage.ImageToPictureDisp(Properties.Resources.icon_icMaintenanceTools_s);
                    stdole.IPictureDisp oImageLarge = ConvertImage.ImageToPictureDisp(Properties.Resources.icon_icMaintenanceTools_l);
                    m_buttonForm = piAddinSite.CreateCommandHandler("icMaintenanceTools", "icMaintenanceTools", "icMaintenanceTools", "IRONCADの修復や設定変更を行うツールです。", oImageSmall, oImageLarge);
                    m_buttonForm.Enabled = true;

                    //Control bar
                    ZControlBar cControlBar;
                    ZEnvironmentMgr cEnvMgr = this.IronCADApp.EnvironmentMgr;
                    ZControls cControls;
                    IZControl cControl;
                    ZRibbonBar cRibbonBar;

                    //ツールバーを作成する
                    IZEnvironment cEnv = cEnvMgr.get_Environment(eZEnvType.Z_ENV_SCENE);
                    cRibbonBar = cEnv.GetRibbonBar(eZRibbonBarType.Z_RIBBONBAR);
                    cControlBar = cEnv.AddControlBar(piAddinSite, "icAPI_Sample_C#_ControlBar");
                    cControls = cControlBar.Controls;
                    cControl = cControls.Add(ezControlType.Z_CONTROL_BUTTON, m_buttonForm.ControlDescriptor, null);

                    //Add button to RibbonBar
                    cRibbonBar.AddButton2(m_buttonForm.ControlDescriptor, true);
//                    cRibbonBar.AddButton2(m_buttonForm.ControlDescriptor, false);

                    /************************************************************
                      リボンバーに大きいボタンで表示させたい時はこっち↓を使用する
                      cRibbonBar.AddButton2(m_button.ControlDescriptor, true);
                    *************************************************************/


                    //Event handlers
                    m_buttonForm.OnClick += new _IZCommandEvents_OnClickEventHandler(m_buttonForm_OnClick);
                    m_buttonForm.OnUpdate += new _IZCommandEvents_OnUpdateEventHandler(m_buttonForm_OnUpdate);

                    //Register App Events
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Addin Server is null.");
            }
        }


        public void DeInitSelf()
        {
            m_buttonForm = null;
         }

#endregion

#region [Private Methods]
        [DllImport("icAPI_CppWrapper.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern IntPtr HwndToCwnd(IntPtr hwnd);

        private void m_buttonForm_OnUpdate()
        {
            m_buttonForm.Enabled = true;  //Change to m_button.Enabled = false; to disable the button  
        }

        private void m_buttonForm_OnClick()
        {
            IZDoc doc = GetActiveDoc();
            IZEnvironmentMgr iZEnvMgr = GetEnvironmentMgr();
            icMaintenanceToolsMain frm = new icMaintenanceToolsMain(this.IronCADApp);
            frm.Show();
        }

        private IZDoc GetActiveDoc()
        {
            if (this.IronCADApp != null)
            {
                return this.IronCADApp.ActiveDoc;
            }
            return null;
        }

        private IZEnvironmentMgr GetEnvironmentMgr()
        {
            if (this.IronCADApp != null)
            {
                return this.IronCADApp.EnvironmentMgr;
            }
            return null;
        }

#endregion

#region [Internal Methods]

        internal static List<IZElement> ConvertObjectToElementArray(object varElements)
        {
            if (varElements != null)
            {
                object[] oElements = varElements as object[];
                if (oElements != null)
                {
                    List<IZElement> izElements = new List<IZElement>();
                    foreach(object oEle in oElements)
                    {
                        IZElement izEle = oEle as IZElement;
                        if (izEle != null)
                        {
                            izElements.Add(izEle);
                        }
                    }
                    return izElements;
                }
            }
            return null;
        }

#endregion

    }
}
