package teste.hssf.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
	
	private static final String nomeDoArquivo = "C:\\wrkspc\\github\\StrutsSample\\planilhaTeste.xlsx";
	
	public static void main(String[] args) {

		leituraDeExcel();

	}

	private static void leituraDeExcel() {
		try {
			
			int contadorLinhas = 1;
			int contadorCelulas = 1;
			
			FileInputStream excelFile = new FileInputStream(new File(nomeDoArquivo));
			
			Workbook workbook = new XSSFWorkbook(excelFile);
			Sheet sheet = workbook.getSheetAt(0);
			Iterator<Row> iteratorLinhas = sheet.iterator();
			
			while(iteratorLinhas.hasNext()) {
				// iterando na linha
				
				Row linhaAtual = iteratorLinhas.next();
				
				Iterator<Cell> iteratorCelulas = linhaAtual.cellIterator();
				
				System.out.println("iteratorLinhas: " + contadorLinhas++);
				
				while(iteratorCelulas.hasNext()) {
					//iterando nas células
					
//					System.out.print("iteratorCelulas: " + contadorCelulas++);
					
					Cell celulaAtual = iteratorCelulas.next();
					
					// verifica tipo do conteúdo da célula
					if(celulaAtual.getCellTypeEnum() == CellType.STRING) {
						System.out.print(celulaAtual.getStringCellValue() + " || ");
					} else if (celulaAtual.getCellTypeEnum() == CellType.NUMERIC) {
						System.out.print(celulaAtual.getNumericCellValue() + " || ");
					}
					
				}
				
				System.out.println();
				
			}
			
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e2) {
			e2.printStackTrace();
		}
	}
}
