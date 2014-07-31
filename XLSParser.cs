using System.Data.Odbc;
using System.Collections.Generic;
using System.Data;

namespace Aquiris.Tools.XLS
{
	public class XLSParser : XLSConnector{
		public string Sheet {set{m_sheet = "["+value + "$]";}}
		private string m_sheet;

		public XLSParser(string p_datasource, string p_sheet = null):base(p_datasource){
			Sheet = p_sheet;
		}

		public DataTable Select(string p_select){
			return ParseQuery("SELECT "+p_select+" FROM "+m_sheet+"");
		}

		private DataTable ParseQuery(string p_selectQuery){
			DataTable result = new DataTable();
			OdbcDataAdapter adapter = new OdbcDataAdapter(p_selectQuery, m_connection);
			adapter.Fill(result);
			return result;
		}
	}
}
