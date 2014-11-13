import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;



//classe eseguibile che presenta un metodo per aggiungere ad un workbook un nuovo foglio xls,
//inserendovi come nuove righe il risultato di una query eseguita su database oracle,
//dove ogni foglio excel contiene i record di una data tabella
public class TestClass {
	
	//metodo che prende in input il workbook su cui aggiungere il foglio, il nome da dare al foglio
	//e l'arraylist di array di stringhe dove ogni array di stringa contiene le celle 
	public void aggiungiSheetPerTabella(HSSFWorkbook wb,String nomeFoglio,ArrayList<String[]> out){
		int nriga,ncol;
		HSSFSheet foglio; 
		nriga=0; 
		foglio=wb.createSheet(nomeFoglio); //ottengo riferimento per un nuovo foglio del workbook
		
		//per ognuno degli arrya di stringhe che si trovano nell'array list
		for(Iterator<String[]> it=out.iterator();it.hasNext();){
			//azzera il contatore di colonne
			ncol=0;
			//ottieni il prossimo array di stringhe (corrispondente quindi ad una nuova riga del foglio excel)
			String[] temp=it.next();
			//crea una nuova riga excel sul foglio, ottenendone il riferimento
			Row riga=foglio.createRow(nriga++);
			//per ogni elemento dell'array di stringhe...
			for (String t : temp){
				//crea e ottieni riferimento alla cella excel della riga sopra creata
				Cell cella=riga.createCell(ncol++);
				//setta come valore della cella, la stringa attuale
				cella.setCellValue(t);
			}
				
		}
	}
	
	
	//la classe è eseguibile
	public static void main(String[] args){
		
		JDBCWrapper wrap1=null;
		ArrayList<String[]> out=null;;
		HSSFWorkbook wb=null;
		
		TestClass obj=new TestClass(); //creo istanza di queste stessa classe poichè il metodo che richiamo non è static
		
		try {
			wrap1=new JDBCWrapper(); //creo oggetto della mia classe wrapper per il jdbc (per oracle)
			wrap1.connetti("system","oracle"); //lancio metodo con user e pw
			wb=new HSSFWorkbook(); //creo nuovo workbook
			
			
			//per ciascuna delle tabelle, relativamente alle quali voglio creare un foglio,
			//uso il metodo ottieniTabellaDaDb (passando nome tabella) dell'oggetto wrapper del jdbc
			//e passo l'arraylist di array di stringhe ritornato, al metodo della classe presente
			//aggiungiSheetPerTabella
			/*-------------------------------------------------------------------*/
			out=wrap1.ottieniTabellaDaDb("schemautente.nometab1");
			if(out!=null)
				obj.aggiungiSheetPerTabella(wb,"IMPIANTI",out);

			/*-------------------------------------------------------------------*/
			out=wrap1.ottieniTabellaDaDb("schemautente.nometab2");
			if(out!=null)
				obj.aggiungiSheetPerTabella(wb,"IMPIEGATI",out);
			
			//creo un FileOutputStream associato al file che voglio scrivere
			FileOutputStream outstream=new FileOutputStream(new File("REPORT_OUTPUT.xls"));
			//lo passo al metodo write del workbook
			wb.write(outstream);
			outstream.close(); //chiudo gli stream
			System.out.println("SCRITTURA EFFETTUATA");
			
			
			
			
			
			if(wrap1!=null)
				wrap1.disconnetti();
			
		} catch (ClassNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		 
		  catch(IOException e){
			  
			 
		  }
		
	}
	
	
	
		 
	
	
}
