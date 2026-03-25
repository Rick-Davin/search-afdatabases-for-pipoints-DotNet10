using OSIsoft.AF.Asset;
using OSIsoft.AF.PI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Search_AFDatabases_for_PIPoints.Models
{
    internal interface IReportWriter
    {
        void CustomDispose();
        void CustomSave();
        void Initialize();
        void Save(bool disposing = false);

        void WriteRows(IDictionary<AFAttribute, PIPoint> map,
                       PIPointList tags,
                       IDictionary<PIPoint, AFValue> firstRecordedDict,
                       IDictionary<PIPoint, AFValue> currentValueDict);

    }
}
