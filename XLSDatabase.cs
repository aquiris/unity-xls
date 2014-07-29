using System.Data.Odbc;
using System.Collections.Generic;
using System.Data;

//remove
using UnityEngine;

namespace Aquiris.Tools.XLS
{
	public class XLSDatabase : XLSConnector{
		public string Sheet {set{m_sheet = "["+value + "$]";}}
		private string m_sheet;

		public XLSDatabase(string p_datasource, string p_sheet) : base(p_datasource){
			Sheet = p_sheet;
		}

		public IDataReader Select(string p_column, string p_where){
			return ExecuteQuery("SELECT "+p_column+" FROM "+m_sheet + " WHERE " + p_where);
		}
		
		public void Insert(Dictionary<string, string> p_entries){
			string keys = "";
			string values = "";
			foreach(KeyValuePair<string,string> pair in p_entries){
				if(!string.IsNullOrEmpty(keys)){
					keys += ",";
					values += ",";
				}
				keys += pair.Key;
				values += "'" + pair.Value + "'";
			}
			ExecuteNonQuery("INSERT INTO "+m_sheet+" (" + keys + ") VALUES (" + values + ");");
		}


	}
}
