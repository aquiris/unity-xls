using System.Data.Odbc;
using System.Collections.Generic;
using System.Data;

using UnityEngine;

namespace Aquiris.Tools.XLS
{
	public class XLSParser : XLSConnector{
		public string Sheet {set{m_sheet = "["+value + "$]";}}
		private string m_sheet;

		public XLSParser(string p_datasource, string p_sheet = null):base(p_datasource){
			Sheet = p_sheet;
		}

		public string GetCell(string p_column, string p_where){
			DataTable table = ParseQuery("SELECT "+p_column+" FROM "+m_sheet+" WHERE "+p_where);
			return table.Rows[0][p_column].ToString();
		}

		public List<string> GetColumn(string p_column){
			List<string> result = new List<string>();
			DataTable table = ParseQuery("SELECT "+p_column+" FROM "+m_sheet+"");
			foreach(DataRow row in table.Rows){
				string s = row[0].ToString();
				result.Add(s);
			}
			return result;
		}

		public DataTable ParseQuery(string p_selectQuery){
			DataTable result = new DataTable();
			OdbcDataAdapter adapter = new OdbcDataAdapter(p_selectQuery, m_connection);
			adapter.Fill(result);
			return result;
		}
	}
}
