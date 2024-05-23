package com.ty;

import java.io.IOException;
import java.io.PrintWriter;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;

import jakarta.servlet.GenericServlet;
import jakarta.servlet.ServletException;
import jakarta.servlet.ServletRequest;
import jakarta.servlet.ServletResponse;

public class Fetch extends javax.servlet.GenericServlet{

	

	@Override
	public void service(javax.servlet.ServletRequest req, javax.servlet.ServletResponse resp)
			throws javax.servlet.ServletException, IOException {
int id = Integer.parseInt(req.getParameter("id"));
		
		//JDBC
		
		String url = "jdbc:mysql://localhost:3306?user=root&password=12345678";
		
		String fqcn = "com.mysql.jdbc.Driver";
		
		String qry = "Select * from employee.info where id = ? ";
		
		PrintWriter out = resp.getWriter();
		
		try {
			
			Class.forName(fqcn);
			
			Connection con = DriverManager.getConnection(url);
			
			PreparedStatement pstmt = con.prepareStatement(qry);
			
			pstmt.setInt(1, id);
			ResultSet rs = pstmt.executeQuery();
			
			
			while(rs.next())
			{
//				System.out.println(rs.getInt(1)+" "+rs.getString(2)+" "+rs.getString(3)+" "+rs.getString(4));
				out.print("<html>"+"<body style='background-color: aquamarine;'>"+"<h1> EmpId :"+rs.getInt(1)+"</h1>"+"<h1>Ename : "+rs.getString(2)+"</h1>"+"<h1> Salary : "+rs.getString(3)+"</h1>"+"<h1> Desgination : "+rs.getString(4)+"</h1>"+"</body>"+"</html>");
			}
			
			
		} catch (Exception e) {
			
			e.printStackTrace();
		}
		
		
	}

}
