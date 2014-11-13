

import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.Locale;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;

public class JDBCWrapper {
	 
	
	//per oracle in remoto
	private static final String driverurl="oracle.jdbc.OracleDriver";
	private static final String dburl="jdbc:oracle:thin:@//iphost:porta/nomeservizio";
	
	private Connection conn1;
	
	public JDBCWrapper() throws ClassNotFoundException{
	 
		Class.forName(driverurl);
		
	}
	
	public void connetti(String user,String pw) throws SQLException{
		conn1=DriverManager.getConnection(dburl,user,pw);
		System.out.println("CONNESSO!");
		
	}
	
	public void disconnetti() throws SQLException{
		if(conn1!=null)
			conn1.close();	
	}
	
	
	public ArrayList<String[]> ottieniTabellaDaDb(String nomeTab) throws SQLException{
		if(conn1==null)
				return null;
		
		ArrayList<String[]> lista=new ArrayList();
		
		Statement st=conn1.createStatement(); //il prepared statement non accetta che si imposti il nome tabella, quindi la concateniamo noi
			
		ResultSet rs=st.executeQuery("select * from "+nomeTab);
		ResultSetMetaData rsmd=rs.getMetaData();
		
		System.out.println(nomeTab + String.valueOf(rsmd.getColumnCount()));
		
		String[] header=new String[rsmd.getColumnCount()];
		
		
		for(int i=0;i<rsmd.getColumnCount();i++)
			header[i]=new String(rsmd.getColumnName(i+1));
		
		lista.add(header);
		
		while(rs.next()){
			String[] riga=new String[rsmd.getColumnCount()];
			
			for(int i=0;i<rsmd.getColumnCount();i++){
				
				String t=rs.getString(i+1);
				
				if(t==null)
					t="NULL";
				
				riga[i]=new String(t);
			}
			
			lista.add(riga);
		}
		
		return lista;
		
	
	}
	
	
	

	
}
