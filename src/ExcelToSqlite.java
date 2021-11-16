import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.*;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.sql.*;
import org.sqlite.JDBC;
import org.sqlite.SQLiteConfig;

public class ExcelToSqlite {
    
    String dbName = "db.db";
    
    public static void main(String[] args) throws SQLException {
        if(args.length > 0 && args[0].endsWith(".xlsx")) {
        	ExcelToSqlite a = new ExcelToSqlite();
	        if(args.length > 1  && args[1].endsWith(".db"))
	        	a.dbName = args[1];
	        a.importFromExcel(args[0]);
        }else {
        	System.out.println("expected excel file path (.xlsx extension) and database file path (.db extension).");
        }
    }

    private Connection connect(String dbPath) throws SQLException {
    	DriverManager.registerDriver(new JDBC());
        String url = "jdbc:sqlite:" + dbPath;
        SQLiteConfig config = new SQLiteConfig();
        config.enforceForeignKeys(true);
        try{
            return DriverManager.getConnection(url, config.toProperties());
        }catch(SQLException e){
            System.out.println(e.getMessage());
        }
        return null;
    }

    public void importFromExcel(String xlPath) throws SQLException {
//    	System.setErr(null);
    	try {
			System.setErr(new PrintStream(new FileOutputStream("err.txt")));
		}
        catch (FileNotFoundException e1) {
			e1.printStackTrace();
		}
        try (Connection con = connect(dbName); XSSFWorkbook wb = new XSSFWorkbook(new File(xlPath));){
			for(int i = 0; i < wb.getNumberOfSheets(); i++) {
				XSSFSheet sheet = wb.getSheetAt(i);
				int lastRow = sheet.getLastRowNum();
				XSSFRow row = sheet.getRow(0);
				short lastCol = row.getLastCellNum();
				// create table query
				Statement stat = con.createStatement();
				String sql = "create table if not exists " + sheet.getSheetName() + "(";
				for(short colNum = 0; colNum < lastCol; colNum++) {
					// the column header required to hold its name, type and constrains (not null/primary key/check statement etc.)
					sql += row.getCell(colNum).getStringCellValue() + ",";
				}
				sql = sql.substring(0, sql.length()-1) + ");";
				stat.execute(sql);
				stat.close();
				for (short rowNum = 1; rowNum < lastRow; rowNum++) {
					row = sheet.getRow(rowNum);
					String rowValues = "";
					for(short colNum = 0; colNum < lastCol; colNum++) {
						XSSFCell cell = row.getCell(colNum);
						if(cell != null && !cell.toString().equals(""))
							rowValues += "\"" + cell.toString() + "\",";// ="<value>",
						else rowValues += "\"\",";// ="",					
					}
					// insert query
					stat = con.createStatement();
					sql = "insert into " + sheet.getSheetName() + " values(" + rowValues.substring(0, rowValues.length()-1) + ");";
					stat.execute(sql);
					stat.close();
				}
			}
		} catch (InvalidFormatException e) {
			System.out.println("an error occured. please check \"err.txt\" for additional information.");
			e.printStackTrace(System.err);
		} catch (IOException e) {
			System.out.println("an error occured. please check \"err.txt\" for additional information.");
			e.printStackTrace(System.err);
		}
    }
}