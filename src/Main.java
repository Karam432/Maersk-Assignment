import java.io.FileOutputStream;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	public static void main(String[] args) throws Exception {
		//list to store fibonacci series upto 20
		ArrayList<Integer> list = new ArrayList<>();
		int prev =0,next=1;
		list.add(next);
		while(next<=20) {
			int temp = next+prev;
			if(temp>20)break;
			list.add(temp);
			prev = next;
			next = temp;
		}
		
		//workbook to create fibonacci series
		XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
		XSSFSheet xssfSheet = xssfWorkbook.createSheet("Sheet");
		XSSFRow xssfRow = xssfSheet.createRow(0);
		xssfRow.createCell(0).setCellValue("Fibo Series");

		XSSFCellStyle styleLast = xssfWorkbook.createCellStyle();
		styleLast.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		styleLast.setFillForegroundColor(IndexedColors.RED.getIndex());
		
		//adding fibonacci series from arraylist to rows
		for(int i=list.size()-1; i>=0; i--) {
			Integer val = list.size()-i;
			XSSFRow tempRow = xssfSheet.createRow(val);
			Integer tempStorage = list.get(i);
			tempRow.createCell(0).setCellValue((Integer)tempStorage);
			if(tempStorage%2 !=0) {
				xssfSheet.getRow(val).getCell(0).setCellStyle(styleLast);
			}
		}
		
		//create output file of fibonacci series
		String outputPathString = ".\\output.xlsx";
		FileOutputStream outputStream = new FileOutputStream(outputPathString);
		xssfWorkbook.write(outputStream);
		outputStream.close();
	}

}
