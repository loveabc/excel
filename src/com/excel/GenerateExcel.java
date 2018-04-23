import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import cn.cathaylife.DataSourceConnection;

/**
 * 
 * @author  生成EXCEL
 *
 */
public class GenerateExcel {

	private static final String excelName = "E://分摊比例.xls";

	public static void generateExcel(List<String> list) throws IOException {
		HSSFWorkbook book = new HSSFWorkbook();
		HSSFSheet sheet = book.createSheet("分摊比例");
		HSSFRow row = sheet.createRow(0);
		HSSFCell cell0 = row.createCell(0);
		HSSFCell cell1 = row.createCell(1);
		HSSFCell cell2 = row.createCell(2);
		HSSFCell cell3 = row.createCell(3);
		HSSFCell cell4 = row.createCell(4);
		HSSFCell cell5 = row.createCell(5);
		HSSFCell cell6 = row.createCell(6);
		cell0.setCellValue("年份");
		cell1.setCellValue("编列单位");
		cell2.setCellValue("预算科目");
		cell3.setCellValue("费率计算系数");
		cell4.setCellValue("渠道分摊系数");
		cell5.setCellValue("导入时间");
		cell6.setCellValue("导入人工号");

		for (int i = 0; i < list.size(); i++) {
			String[] info = list.get(i).split("#");
			row = sheet.createRow(i + 1);
			cell0 = row.createCell(0);
			cell1 = row.createCell(1);
			cell2 = row.createCell(2);
			cell3 = row.createCell(3);
			cell4 = row.createCell(4);
			cell5 = row.createCell(5);
			cell6 = row.createCell(6);
			cell0.setCellValue(info[0]);
			cell1.setCellValue(info[1]);
			cell2.setCellValue(info[2]);
			cell3.setCellValue(info[3]);
			cell4.setCellValue(info[4]);
			cell5.setCellValue(info[5]);
			cell6.setCellValue(info[6]);
			System.out.println(i);
		}

		OutputStream os = new FileOutputStream(new File(excelName));
		try {
			book.write(os);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			os.close();
		}

	}

	public static List<String> readDB() throws ClassNotFoundException, IOException, SQLException {
		Connection conn = new DataSourceConnection().getDataSouceConnection();

		String querySQL = "select cost_yy,pln_divno,cost_cd,index,v_chl_ind,tns_time,tns_empno from dbdk.dtdkh906 where cost_yy='2018'";

		PreparedStatement queryPs = conn.prepareStatement(querySQL);
		ResultSet rs = queryPs.executeQuery();

		List<String> infos = new ArrayList<>();
		while (rs.next()) {
			String info = rs.getString(1) + "#" + rs.getString(2) + "#" + rs.getString(3) + "#" + rs.getString(4) + "#"
					+ rs.getString(5) + "#" + rs.getString(6) + "#" + rs.getString(7) ;
			infos.add(info);
		}
		return infos;

	}

	public static void main(String[] args) throws ClassNotFoundException, SQLException {
		try {
			GenerateExcel.generateExcel(GenerateExcel.readDB());
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
