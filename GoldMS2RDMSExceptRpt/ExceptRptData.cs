using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using Oracle.DataAccess.Client;
using System.Data;
using System.Data.Common;

namespace GoldMS2RDMSExceptRpt
{
    class ExceptRptData
    {
        private DataSet ds;
        public ExceptRptData(OracleConnection conn)
        {
            OracleDataAdapter da = new OracleDataAdapter("SELECT a.*, b.sales_order_number, b.sales_order_line_item, b.stockroom, b.quantity_shipped, b.ship_date, b.shipper_status, b.shipper_sequence_number, b.bill_of_lading_number FROM (SELECT c.* FROM msms.msms_system_compare_h_vw c) a, msms.msms_rdms_shipper_mvw b WHERE (a.gold_due <> a.ms2_due OR a.ms2_due <> a.rdms_due OR a.rdms_due <> a.gold_due) AND a.cap_order = b.cap_order_number(+) AND a.rdms_rcon = b.received_order_number(+) ORDER BY a.part, a.cap_order", conn);
            ds = new DataSet();
            da.Fill(ds, "ExceptRpt");
        }

        public DataSet  DS 
        {
            get { return ds; }
        }

    }
}
