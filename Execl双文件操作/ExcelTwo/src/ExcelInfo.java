

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
	 * 读取一个excel文件的内容
	 * 
	 * @param args
	 * @throws Exception
	 */
	private static final File file = new File("C:\\Users\\我是爸爸妈妈的小甜心\\Desktop\\123\\附件3全口径劳动力台账.xls");

	public static void main(String[] args) throws Exception {
		// extractor();
		readTable();
	}
	
	
	// 通过对单元格遍历的形式来获取信息 ，这里要判断单元格的类型才可以取出值
	public static void readTable() throws Exception {
		// 上面姓名
		InputStream ips = new FileInputStream(
				"C:\\Users\\我是爸爸妈妈的小甜心\\Desktop\\123\\附件3全口径劳动力台账.xls");
		HSSFWorkbook wb = new HSSFWorkbook(ips);
		HSSFSheet sheet = wb.getSheetAt(2);
		// 下面身份证
		InputStream ips1 = new FileInputStream(
				"C:\\Users\\我是爸爸妈妈的小甜心\\Desktop\\123\\响水社区(4).xls");
		HSSFWorkbook wb1 = new HSSFWorkbook(ips1);
		HSSFSheet sheet1 = wb1.getSheetAt(0);
		for (int i = 6; i < sheet.getPhysicalNumberOfRows(); i++) {
			HSSFRow row = sheet.getRow(i);
			String userName = String.valueOf(row.getCell(7));// 需要填写表格的用户名
			for (int j = 0; j < sheet1.getPhysicalNumberOfRows(); j++) {
				HSSFRow row1 = sheet1.getRow(j);
				String userName1 = String.valueOf(row1.getCell(2));// 需要查询的表格的用户名
				if (userName.equals(userName1) && String.valueOf(row1.getCell(3)).equals("户主")) {
					ExcelInfo.updateExcel(file, 2, 8, i,String.valueOf(row1.getCell(7)));
				} else {
					continue;
				}
			}
			System.out.println(i);
		}
		System.out.println("运行结束！");
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
            System.out.println("您是否打开了excel文件！请关闭后再试。");
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
 
        fis.close();//关闭文件输入流
 
        FileOutputStream fos=new FileOutputStream(exlFile);
        try {
        	workbook.write(fos);
		} catch (Exception e) {
			e.printStackTrace();
		}
        fos.close();//关闭文件输出流
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
