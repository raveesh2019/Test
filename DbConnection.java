package com.fillo;


import java.sql.Connection;
import java.sql.DriverManager;

public class DbConnection {

	
	public static Connection getConnection() {
		// TODO Auto-generated method stub
	
		Connection conn=null;
		try {
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
			conn = DriverManager.getConnection("jdbc:sqlserver://TTWDERBI01:1435;databaseName=ERBI_UIU;integratedSecurity=true");
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return conn;
	}
}
