using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swpublished;
using System;
using System.Runtime.InteropServices;

namespace SACO.SolidWorks.ComptProps
{
    public class PropsAddinMain : SwAddin
    {

        #region Public Memebers

        public const string SWTASKPANE_PROGID = "Component.Properties.New";

        #endregion

        #region Private Members

        private int mSwCookie;

        private TaskpaneView mTaskpaneView;

        private TaskpaneHost mTaskPaneHost;

        static ISldWorks mSolidWorksApplication;



        private ModelDoc2 swModel;
        private ModelDocExtension modelExt;
        private CustomPropertyManager customProps;


        #endregion

        #region Solidworks Add-in Callbacks

        public bool ConnectToSW(object ThisSW, int Cookie)
        {
            mSolidWorksApplication = (ISldWorks)ThisSW;

            mSwCookie = Cookie;

            var ok = mSolidWorksApplication.SetAddinCallbackInfo2(0, this, mSwCookie);

            var val = "";
            var valOut = "";

            swModel = mSolidWorksApplication.ActiveDoc;

            if (swModel != null) // && getExtType == (int)swDocumentTypes_e.swDocPART)
            {
                var getExtType = swModel.GetType();
                modelExt = swModel.Extension;
                customProps = modelExt.CustomPropertyManager[""];
                var tempType = customProps.Get4("Template_Type", false, out val, out valOut);

                if (tempType == true)
                {
                    LoadUI();
                }
                else
                {
                    UnloadUI();
                }
            }
            else
            {
                LoadUI();
            }


            

            return true;
        }


        public bool DisconnectFromSW()
        {
            UnloadUI();

            return true;
        }

        #endregion

        #region Create UI

        private void LoadUI()
        {
            mTaskpaneView = mSolidWorksApplication.CreateTaskpaneView2(string.Empty, "This is a task pane");

            mTaskPaneHost = (TaskpaneHost)mTaskpaneView.AddControl(PropsAddinMain.SWTASKPANE_PROGID, string.Empty);

            mTaskPaneHost.getSwApp((SldWorks)mSolidWorksApplication);//added
        }

        private void UnloadUI()
        {
            mTaskPaneHost = null;

            mTaskpaneView.DeleteView();

            Marshal.ReleaseComObject(mTaskpaneView);

            mTaskpaneView = null;
        }

        #endregion

        #region COM Registration

        [ComRegisterFunction()]
        private static void ComRegister(Type t)
        {
            var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

            using (var rk = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(keyPath))
            {
                rk.SetValue(null, 1);

                rk.SetValue("Title", "Component Properties");
                rk.SetValue("Description", "Set and get the properties for a component.");
            }
        }

        [ComUnregisterFunction()]
        private static void ComUnregister(Type t)
        {
            var keyPath = string.Format(@"SOFTWARE\SolidWorks\AddIns\{0:b}", t.GUID);

            Microsoft.Win32.Registry.LocalMachine.DeleteSubKey(keyPath);
        }

        #endregion
    }


   
    
}
