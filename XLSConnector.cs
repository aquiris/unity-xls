using System.Data.Odbc;
using System.Collections.Generic;
using System.Data;

//remove
using UnityEngine;

namespace Aquiris.Tools.XLS
{
	public class XLSConnector{
		protected OdbcConnection m_connection;

		private OdbcCommand m_command;
		private OdbcDataReader m_reader;

		//carefull here, it is fucky stuff
		private string m_connectionString = 
			"Driver={Microsoft Excel Driver (*.xls)};" +
			"DriverId=790;" +
			"UNICODESQL=1;" +
			"Unicode=yes;" +
			"AllowFormula=false;" +
			"READONLY=0;";

		public XLSConnector(string p_datasource){
			m_connection = new OdbcConnection();
			SetDataSource(p_datasource);
		}

		public void SetDataSource(string p_datasource){
			m_connection.ConnectionString = m_connectionString + "DBQ="+p_datasource+";";
			m_command = m_connection.CreateCommand();
		}

		protected void ExecuteNonQuery(string p_nonquery){
			m_command.CommandText = p_nonquery;
			try{
				m_connection.Open();
				m_command.ExecuteNonQuery();
			}catch(DataException ex){
				Debug.LogError(ex.Message);
			}finally{
				m_connection.Close();
			}
		}

		protected IDataReader ExecuteQuery(string p_query){
			m_command.CommandText = p_query;
			try{
				m_connection.Open();
				m_reader = m_command.ExecuteReader();
			}catch(DataException ex){
				Debug.LogError(ex.Message);
			}finally{
				m_connection.Close();
			}
			return m_reader;
		}
	}
}
