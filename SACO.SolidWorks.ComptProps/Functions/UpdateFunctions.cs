using SACO.SolidWorks.ComptProps.Models;
using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SACO.SolidWorks.ComptProps.Functions
{
    public class UpdateFunctions
    {
        private SldWorks swApp;
        private ModelDoc2 swModel;
        private ModelDocExtension modelExt;
        private CustomPropertyManager customProps;

        SacoComponentPropertiesModel propsModel;

        public UpdateFunctions(SldWorks SwApp, SacoComponentPropertiesModel PropsModel)
        {
            swApp = SwApp;
            swModel = swApp.ActiveDoc;
            propsModel = PropsModel;
        }

        public void SaveProps()
        {
            modelExt = swModel.Extension;
            customProps = modelExt.CustomPropertyManager[""];

            //loop through each property in propsmodel and set to the required property by doing a
            foreach(var propItem in propsModel.GetType().GetProperties())
            {
                customProps.Set2(propItem.Name, propItem.GetValue(propItem, null).ToString());
            }

        }
    }
}
