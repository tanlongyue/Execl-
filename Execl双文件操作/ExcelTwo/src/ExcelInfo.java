

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.CellRange;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class ExcelInfo {
	/**
	 * ��ȡһ��excel�ļ�������
	 * 
	 * @param args
	 * @throws Exception
	 */
	private static final File file = new File("C:\\Users\\���ǰְ������С����\\Desktop\\123\\����3ȫ�ھ��Ͷ���̨��.xls");

	public static void main(String[] args) throws Exception {
		// extractor();
		readTable();
	}
	
	
	// ͨ���Ե�Ԫ���������ʽ����ȡ��Ϣ ������Ҫ�жϵ�Ԫ������Ͳſ���ȡ��ֵ
	public static void readTable() throws Exception {
		// ��������
		InputStream ips = new FileInputStream(
				"C:\\Users\\���ǰְ������С����\\Desktop\\123\\����3ȫ�ھ��Ͷ���̨��.xls");
		HSSFWorkbook wb = new HSSFWorkbook(ips);
		HSSFSheet sheet = wb.getSheetAt(2);
		// �������֤
		InputStream ips1 = new FileInputStream(
				"C:\\Users\\���ǰְ������С����\\Desktop\\123\\��ˮ����(4).xls");
		HSSFWorkbook wb1 = new HSSFWorkbook(ips1);
		HSSFSheet sheet1 = wb1.getSheetAt(0);
		for (int i = 6; i < sheet.getPhysicalNumberOfRows(); i++) {
			HSSFRow row = sheet.getRow(i);
			String userName = String.valueOf(row.getCell(7));// ��Ҫ��д�����û���
			for (int j = 0; j < sheet1.getPhysicalNumberOfRows(); j++) {
				HSSFRow row1 = sheet1.getRow(j);
				String userName1 = String.valueOf(row1.getCell(2));// ��Ҫ��ѯ�ı����û���
				if (userName.equals(userName1) && String.valueOf(row1.getCell(3)).equals("����")) {
					ExcelInfo.updateExcel(file, 2, 8, i,String.valueOf(row1.getCell(7)));
				} else {
					continue;
				}
			}
			System.out.println(i);
		}
		System.out.println("���н�����");
		return;
	}
    public static void modifyExcel(int SheetNo, int lieshu, int hangshu, String value) {
        try {
            Workbook rwb = Workbook.getWorkbook(file);
            WritableWorkbook wwb = Workbook.createWorkbook(file, rwb);// copy
            WritableSheet ws = wwb.getSheet(SheetNo);
            Label labelCF = new Label(lieshu, hangshu, value);
            ws.addCell(labelCF);
            wwb.write();
            wwb.close();
            rwb.close();
        } catch (Exception e) {
            e.printStackTrace();
            System.out.println("���Ƿ����excel�ļ�����رպ����ԡ�");
        }
    }
    
    public static void updateExcel(File exlFile,int sheetIndex,int col,int row,String value)throws Exception{
        FileInputStream fis=new FileInputStream(exlFile);
        HSSFWorkbook workbook=new HSSFWorkbook(fis);
//        workbook.
        HSSFSheet sheet=workbook.getSheetAt(sheetIndex);
        
        HSSFRow r=sheet.getRow(row);
        HSSFCell cell=r.getCell(col);
//        int type=cell.getCellType();
//        String str1=cell.getStringCellValue();
        cell.setCellValue(value);
 
        fis.close();//�ر��ļ�������
 
        FileOutputStream fos=new FileOutputStream(exlFile);
        try {
        	workbook.write(fos);
		} catch (Exception e) {
			e.printStackTrace();
		}
        fos.close();//�ر��ļ������
    }
 
 
    private String getCellValue(HSSFCell cell) {
        String cellValue = "";
        DecimalFormat df = new DecimalFormat("#");
        switch (cell.getCellType()) {
            case XSSFCell.CELL_TYPE_STRING:
                cellValue = cell.getRichStringCellValue().getString().trim();
                break;
            case XSSFCell.CELL_TYPE_NUMERIC:
                cellValue = df.format(cell.getNumericCellValue()).toString();
                break;
            case XSSFCell.CELL_TYPE_BOOLEAN:
                cellValue = String.valueOf(cell.getBooleanCellValue()).trim();
                break;
            case XSSFCell.CELL_TYPE_FORMULA:
                cellValue = cell.getCellFormula();
                break;
            default:
                cellValue = "";
        }
        return cellValue;
    }
 
}
