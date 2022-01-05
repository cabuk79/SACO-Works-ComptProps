using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SACO.SolidWorks.ComptProps.Models
{
    public class SacoComponentPropertiesModel
    {

        #region General

        public string UncontrolledRevision { get; set; }
        public string ControlledRevision { get; set; }
        public string Thickness { get; set; }
        public string MaterialInfo { get; set; }
        public string Project { get; set; }
        public string FinishInfo { get; set; }
        public string TitleInfo { get; set; }
        public string CommentsNotes { get; set; }
        public string TeamplteType { get; set; }
        public string ProductGroup { get; set; }

        #endregion

        #region Uncontrolled

        public string DrawnUncontolled { get; set; }
        public DateTime DrawnUncontrolledByDate { get; set; }
        public string UncontrolledCheckedByOne { get; set; }
        public string UncontrolledCheckedByTwo { get; set; }
        public string UncontrolledCheckedByThree { get; set; }
        public DateTime UncontrolledCheckedByDateOne { get; set; }
        public DateTime UncontrolledCheckedByDateTwo { get; set; }
        public DateTime UncontrolledCheckedByDateThree { get; set; }

        #endregion

        #region Controlled

        public string DrawnContolled { get; set; }
        public DateTime DrawnControlledByDate { get; set; }
        public string CntrolledCheckedByOne { get; set; }
        public string ControlledCheckedByTwo { get; set; }
        public string ControlledCheckedByThree { get; set; }
        public DateTime ControlledCheckedByDateOne { get; set; }
        public DateTime ControlledCheckedByDateTwo { get; set; }
        public DateTime ControlledCheckedByDateThree { get; set; }
        public string SignedBy { get; set; }
        public DateTime SignedDate { get; set; }

        #endregion

    }
}
