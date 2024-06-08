package Excel.Ex;

import java.io.*;
import java.sql.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

public class Excel2DatabaseTest {

    public static void main(String[] args) {
        String jdbcURL = "jdbc:mysql://localhost:3306/servlet?useSSL=false";
        String username = "root";
        String password = "root";
        String excelFilePath = "C:\\Users\\GANESH\\Downloads\\spdclmasterdata.xlsx";
        int batchSize = 20;
        Connection connection = null;

        try {
            long start = System.currentTimeMillis();
            FileInputStream inputStream = new FileInputStream(excelFilePath);
            Workbook workbook = new XSSFWorkbook(inputStream);
            Sheet firstSheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = firstSheet.iterator();
            connection = DriverManager.getConnection(jdbcURL, username, password);
            connection.setAutoCommit(false);

            String sql = "INSERT INTO excel (CIRCD, CIRNAME, DIVCD, DIVNAME, EROCD, ERONAME, SUBCD, SUBNAME, SECCD, SECNAME) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
            PreparedStatement statement = connection.prepareStatement(sql);    

            int count = 0;

            rowIterator.next();

            while (rowIterator.hasNext()) {
                Row nextRow = rowIterator.next();
                Iterator<Cell> cellIterator = nextRow.cellIterator();

                while (cellIterator.hasNext()) {
                    Cell nextCell = cellIterator.next();
                    int columnIndex = nextCell.getColumnIndex();

                    switch (columnIndex) {
                        case 0:
                            String CIRCD = nextCell.getStringCellValue();
                            statement.setString(1, CIRCD);
                            break;
                        case 1:
                            String CIRNAME = nextCell.getStringCellValue();
                            statement.setString(2, CIRNAME);
                            break;
                        case 2:
                            String DIVCD = nextCell.getStringCellValue();
                            statement.setString(3, DIVCD);
                            break;
                        case 3:
                            String DIVNAME = nextCell.getStringCellValue();
                            statement.setString(4, DIVNAME);
                            break;
                        case 4:
                            String EROCD = nextCell.getStringCellValue();
                            statement.setString(5, EROCD);
                            break;
                        case 5:
                            String ERONAME = nextCell.getStringCellValue();
                            statement.setString(6, ERONAME);
                            break;
                        case 6:
                            String SUBCD = nextCell.getStringCellValue();
                            statement.setString(7, SUBCD);
                            break;
                        case 7:
                            String SUBNAME = nextCell.getStringCellValue();
                            statement.setString(8, SUBNAME);
                            break;
                        case 8:
                            String SECCD = nextCell.getStringCellValue();
                            statement.setString(9, SECCD);
                            break;
                        case 9:
                            String SECNAME = nextCell.getStringCellValue();
                            statement.setString(10, SECNAME);
                            break;
                    }   
                }

                statement.addBatch();

                if (++count % batchSize == 0) {
                    statement.executeBatch();
                }              
            }

            workbook.close();

            statement.executeBatch();
            connection.commit();
            connection.close();

            long end = System.currentTimeMillis();
            System.out.printf("Import done in %d ms\n", (end - start));

        } catch (IOException ex1) {
            System.out.println("Error reading file");
            ex1.printStackTrace();
        } catch (SQLException ex2) {
            System.out.println("Database error");
            ex2.printStackTrace();
        }
    }
}
