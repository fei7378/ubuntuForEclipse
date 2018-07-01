package test;

import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

public class createBlankExcel {
	public static void main(String[] args) {
		List<AccountStatementExcel> users = new ArrayList<>();
		AccountStatementExcel user1 = new AccountStatementExcel("12", "12", "12", "12", 12, 13, 14, 15, "12",
				"12", "12", "12", "1112", "112");
		AccountStatementExcel user2 = new AccountStatementExcel("12", "12", "12", "12", 12, 13, 14, 15, "12",
				"12", "12", "12", "12222", "122");
		AccountStatementExcel user3 = new AccountStatementExcel("12", "12", "12", "12", 12, 13, 14, 15, "12",
				"12", "12", "12", "1332", "132");
		users.add(user1);
		users.add(user2);
		users.add(user3);
		System.out.println(users.size());
		
		String[] first = datatime();
		try {
			HSSFWorkbook workbook = new HSSFWorkbook();
			workbook = getHSSFWorkbook(workbook, first, null, users);
			FileOutputStream out = new FileOutputStream("C:/Users/飞云川/Desktop/对账单test.xls");
			workbook.write(out);
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		try {
			Runtime.getRuntime().exec("cmd /c start C:/Users/飞云川/Desktop/对账单test.xls");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	/**
	 * 导出Excel
	 * 
	 * @param sheetName
	 *            sheet名称
	 * @param title
	 *            标题
	 * @param values
	 *            内容
	 * @param wb
	 *            HSSFWorkbook对象
	 * @return
	 */
	@SuppressWarnings("deprecation")
	public static HSSFWorkbook getHSSFWorkbook(HSSFWorkbook wb, String[] first, String[][] Second,
			List<AccountStatementExcel> users) {
		// excel标题
		
		
		String[] titles1 = { "商户编号", "商户名称", "报表生成时间", "所在时区" };
		String[] titles2 = { "全时便利店", "支付渠道", "收款类订单", "退款类订单", "结算" };
		String[] titles3 = { "总金额(元)", "笔数", "应收金额(元)", "交易手续费(元)", "到账金额(元)" };
		String[] titles4 = { "微信支付", "支付宝", "现金支付", "会员支付", "合计" };
		String[] daytime = datatime();
		String[][] Second1 = new String[6][7];
		Second1[0][0] = "全时便利店";

		for (int i = 0; i <= 6; i++) {
			Second1[1][i] = "微信" + i;
		}
		for (int i = 0; i <= 6; i++) {
			Second1[2][i] = "支付宝" + i;
		}
		for (int i = 0; i <= 6; i++) {
			Second1[3][i] = "现金" + i;
		}
		for (int i = 0; i <= 6; i++) {
			Second1[4][i] = "会员" + i;
		}
		for (int i = 0; i <= 6; i++) {
			Second1[5][i] = "合计" + i;
		}
		if (Second == null) {
			Second = Second1;
		}
		// 第一步，创建一个HSSFWorkbook，对应一个Excel文件
		if (wb == null) {
			wb = new HSSFWorkbook();
		}

		// 第二步，在workbook中添加一个sheet,对应Excel文件中的sheet
		HSSFSheet sheet = wb.createSheet("单日汇总");

		sheet.setDefaultRowHeight((short) 315);
		sheet.setDefaultColumnWidth(15);// 13

		// 行号-1

		HSSFRow row0 = sheet.createRow(0);
		HSSFRow row1 = sheet.createRow(1);
		HSSFRow row2 = sheet.createRow(2);
		HSSFRow row3 = sheet.createRow(3);
		HSSFRow row4 = sheet.createRow(4);
		HSSFRow row5 = sheet.createRow(5);
		HSSFRow row6 = sheet.createRow(6);
		HSSFRow row7 = sheet.createRow(7);
		HSSFRow row8 = sheet.createRow(8);
		HSSFRow row9 = sheet.createRow(9);
		HSSFRow row10 = sheet.createRow(10);
		HSSFRow row11 = sheet.createRow(11);
		HSSFRow row12 = sheet.createRow(12);

		// 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制，生成行
		/*
		 * HSSFRow row = sheet.createRow(0); for(int i=0;i<10;i++){ row =
		 * sheet.createRow(i + 1); row.setHeight((short)315); for(int j=0;j<10;j++){
		 * //将内容按顺序赋给对应的列对象 row.createCell(j).setCellValue(j); } }
		 */

		// 第四步，创建单元格样式一，并设置值表头 设置表头居中
		HSSFCellStyle style1 = wb.createCellStyle();
		// 设置这些样式
		style1.setAlignment(HorizontalAlignment.CENTER); // 水平居中
		style1.setVerticalAlignment(VerticalAlignment.CENTER);
		style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);// 设置背景颜色，这句代码必加
		// style1.setFillBackgroundColor(new HSSFColor.LAVENDER().getIndex());
		//
		// style1.setFillForegroundColor(HSSFColor.LIME.index);
		// style1.setFillBackgroundColor(HSSFColor.LAVENDER.index);
		style1.setFillForegroundColor(IndexedColors.CORNFLOWER_BLUE.getIndex());// 设置背景颜色，参考
																				// https://blog.csdn.net/liaomin416100569/article/details/42676681
		// style1.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		String sheetString = (daytime[2] + "交易明细").trim();

		style1.setBorderBottom(BorderStyle.THIN); // 下边框
		style1.setBorderLeft(BorderStyle.THIN);// 左边框
		style1.setBorderTop(BorderStyle.THIN);// 上边框
		style1.setBorderRight(BorderStyle.THIN);// 右边框
		// 设置字体
		HSSFFont boldFont = wb.createFont();
		boldFont.setFontHeightInPoints((short) 10);// 字体大小
		boldFont.setBold(true);
		style1.setFont(boldFont);

		HSSFCellStyle styleColorBotton = wb.createCellStyle();
		styleColorBotton.setFillPattern(FillPatternType.SOLID_FOREGROUND);// 设置背景颜色，这句代码必加
		// styleColorBotton.setFillBackgroundColor(new HSSFColor.BLUE().getIndex());
		// // styleColorBotton.setFillForegroundColor(HSSFColor.LIME.index);
		// styleColorBotton.setFillBackgroundColor(HSSFColor.BLUE.index);
		styleColorBotton.setAlignment(HorizontalAlignment.CENTER); // 水平居中
		styleColorBotton.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
		styleColorBotton.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());// 设置背景颜色，参考https://blog.csdn.net/liaomin416100569/article/details/42676681
		styleColorBotton.setBorderLeft(BorderStyle.THIN);// 左边框
		styleColorBotton.setBorderTop(BorderStyle.THIN);// 上边框
		styleColorBotton.setBorderRight(BorderStyle.THIN);// 右边框
		// styleColorBotton.setBorderBottom(BorderStyle.THICK);// 下边框加粗

		HSSFCellStyle styleTHICKBotton = wb.createCellStyle();
		styleTHICKBotton.setBorderBottom(BorderStyle.THICK);// 下边框加粗

		HSSFCellStyle styleTHICKTop = wb.createCellStyle();
		styleTHICKTop.setBorderTop(BorderStyle.THICK);// 上边框加粗

		HSSFCellStyle styleTHICKRight = wb.createCellStyle();
		styleTHICKRight.setBorderRight(BorderStyle.THICK);// 右边框加粗

		HSSFCellStyle styleTHICKRightAndBotton = wb.createCellStyle();
		styleTHICKRightAndBotton.setBorderBottom(BorderStyle.THICK);// 下边框加粗
		styleTHICKRightAndBotton.setBorderRight(BorderStyle.THICK);// 右边框加粗

		HSSFCellStyle styleTHICKLeftAndBotton = wb.createCellStyle();
		styleTHICKLeftAndBotton.setBorderBottom(BorderStyle.THICK);// 下边框加粗
		styleTHICKLeftAndBotton.setBorderLeft(BorderStyle.THICK);// 右边框加粗

		HSSFCellStyle styleTHICKLeft = wb.createCellStyle();
		styleTHICKLeft.setBorderLeft(BorderStyle.THICK);// 左边框加粗

		HSSFCellStyle styleTHICKLeftAndTHINBotton = wb.createCellStyle();
		styleTHICKLeftAndTHINBotton.setBorderLeft(BorderStyle.THICK);// 左边框加粗
		styleTHICKLeftAndTHINBotton.setBorderBottom(BorderStyle.THIN);// 下边框

		HSSFCellStyle styleTHINBotton = wb.createCellStyle();
		styleTHINBotton.setBorderBottom(BorderStyle.THIN);// 下边框

		HSSFCellStyle styleTHINLeft = wb.createCellStyle();
		styleTHINLeft.setBorderLeft(BorderStyle.THIN);// 下边框

		HSSFCellStyle styleDataFirst = wb.createCellStyle();// 填充上方数据时使用此样式
		styleDataFirst.setBorderLeft(BorderStyle.THIN);// 左边
		styleDataFirst.setAlignment(HorizontalAlignment.CENTER); // 水平居中

		HSSFCellStyle styleDataSecond = wb.createCellStyle();// 填充下方数据时使用此样式
		styleDataSecond.setBorderLeft(BorderStyle.THIN);// 左边框
		styleDataSecond.setBorderTop(BorderStyle.THIN);// 上边框
		styleDataSecond.setBorderRight(BorderStyle.THIN);// 右边框
		styleDataSecond.setBorderBottom(BorderStyle.THIN);// 下边框加粗
		styleDataSecond.setAlignment(HorizontalAlignment.CENTER); // 水平居中
		styleDataSecond.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
		
		
		//************************************第二张表使用样式
		HSSFCellStyle styleTwoFirst = wb.createCellStyle();// 填充列标题与合计
		styleTwoFirst.setBorderLeft(BorderStyle.THIN);// 左边框
		styleTwoFirst.setBorderTop(BorderStyle.THIN);// 上边框
		styleTwoFirst.setBorderRight(BorderStyle.THIN);// 右边框
		styleTwoFirst.setBorderBottom(BorderStyle.THIN);// 下边框
		styleTwoFirst.setAlignment(HorizontalAlignment.CENTER); // 水平居中
		styleTwoFirst.setFillPattern(FillPatternType.SOLID_FOREGROUND);// 设置背景颜色，这句代码必加
		styleTwoFirst.setFillForegroundColor(IndexedColors.TURQUOISE.getIndex());// 设置背景颜色，参考https://blog.csdn.net/liaomin416100569/article/details/42676681
		styleTwoFirst.setAlignment(HorizontalAlignment.CENTER); // 水平居中
		styleTwoFirst.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
		
		
		HSSFCellStyle styleTwoSecond = wb.createCellStyle();// 填充数据
		styleTwoSecond.setBorderLeft(BorderStyle.DOTTED);// 左边框
		styleTwoSecond.setBorderTop(BorderStyle.DOTTED);// 上边框
		styleTwoSecond.setBorderRight(BorderStyle.DOTTED);// 右边框
		styleTwoSecond.setBorderBottom(BorderStyle.DOTTED);// 下边框
		styleTwoSecond.setAlignment(HorizontalAlignment.CENTER); // 水平居中
		styleTwoSecond.setFillPattern(FillPatternType.SOLID_FOREGROUND);// 设置背景颜色，这句代码必加
		styleTwoSecond.setFillForegroundColor(IndexedColors.LIGHT_TURQUOISE.getIndex());// 设置背景颜色，参考https://blog.csdn.net/liaomin416100569/article/details/42676681
		styleTwoSecond.setAlignment(HorizontalAlignment.CENTER); // 水平居中
		styleTwoSecond.setVerticalAlignment(VerticalAlignment.CENTER); // 垂直居中
		// 创建单元格样式二，设置单元格背景色
		
		// HSSFCellStyle style2 = wb.createCellStyle();
		//
		// style2.setFillBackgroundColor(new HSSFColor.GREY_25_PERCENT().getIndex()); //
		// // 设置字体
		// HSSFFont boldFont = wb.createFont();
		// boldFont.setFontHeightInPoints((short)14);//字体大小 boldFont.setBold(true);
		// style2.setFont(boldFont);

		// 将设置合并单元格的样式添加到sheet中
		// 创建标题
		// 声明列对象，生成列
		HSSFCell cell = null;
		/* for(int i=1;i<titles.length;i++){ */
		cell = row1.createCell(0);
		cell.setCellStyle(styleTHICKRight);
		cell = row1.createCell(1);
		cell.setCellValue(titles1[0]);
		cell.setCellStyle(style1);// styleTHICKTop
		// cell.setCellStyle(styleTHICKTop);

		cell = row1.createCell(2);
		cell.setCellStyle(styleTHINBotton);
		cell = row1.createCell(3);
		cell.setCellStyle(styleTHINBotton);

		cell = row1.createCell(4);
		cell.setCellValue(titles1[1]);
		cell.setCellStyle(style1);
		// cell.setCellStyle(styleTHICKTop);
		cell = row1.createCell(5);
		cell.setCellStyle(styleTHINBotton);// styleTHICKTop
		cell = row1.createCell(6);
		cell.setCellStyle(styleTHINBotton);// styleTHICKTop
		cell = row1.createCell(7);
		cell.setCellValue(titles1[2]);
		cell.setCellStyle(style1);
		// cell.setCellStyle(styleTHICKTop);

		cell = row1.createCell(8);
		cell.setCellValue(titles1[3]);
		cell.setCellStyle(style1);
		cell = row1.createCell(9);
		cell.setCellStyle(styleTHINBotton);//styleTwoFirst
		cell = row1.createCell(10);
		cell.setCellStyle(styleTHICKLeft);

		// 下表
		cell = row5.createCell(0);
		cell.setCellStyle(styleTHICKRight);// 右侧边框加粗

		cell = row5.createCell(1);
		cell.setCellValue(titles2[0]);
		cell.setCellStyle(style1);

		cell = row5.createCell(2);
		cell.setCellValue(titles2[1]);
		cell.setCellStyle(style1);

		cell = row5.createCell(3);
		cell.setCellValue(titles2[2]);
		cell.setCellStyle(style1);

		cell = row5.createCell(5);
		cell.setCellValue(titles2[3]);
		cell.setCellStyle(style1);

		cell = row5.createCell(7);
		cell.setCellValue(titles2[4]);
		cell.setCellStyle(style1);

		// 子标题2
		cell = row6.createCell(0);
		cell.setCellStyle(styleTHICKRight);// 右侧边框加粗
		cell = row6.createCell(3);
		cell.setCellValue(titles3[0]);
		cell.setCellStyle(style1);

		cell = row6.createCell(4);
		cell.setCellValue(titles3[1]);
		cell.setCellStyle(style1);

		cell = row6.createCell(5);
		cell.setCellValue(titles3[0]);
		cell.setCellStyle(style1);

		cell = row6.createCell(6);
		cell.setCellValue(titles3[1]);
		cell.setCellStyle(style1);

		cell = row6.createCell(7);
		cell.setCellValue(titles3[2]);
		cell.setCellStyle(style1);

		cell = row6.createCell(8);
		cell.setCellValue(titles3[3]);
		cell.setCellStyle(style1);

		cell = row6.createCell(9);
		cell.setCellValue(titles3[4]);
		cell.setCellStyle(style1);
		// 支付渠道
		cell = row7.createCell(0);
		cell.setCellStyle(styleTHICKRight);// 右侧边框加粗
		cell = row7.createCell(2);
		cell.setCellStyle(styleDataSecond);
		cell.setCellValue(titles4[0]);
		// cell.setCellStyle(style1);

		for (int i = 3; i < 10; i++) {
			cell = row7.createCell(i);
			cell.setCellStyle(styleDataSecond);// 填充第二组数据样式
			cell.setCellValue(Second[1][i - 3]);// 填充第二组数据
		}

		cell = row7.createCell(1);
		cell.setCellStyle(styleDataSecond);// 填充第二组数据样式
		cell.setCellValue(Second[0][0]);// 填充第二组数据

		cell = row8.createCell(0);
		cell.setCellStyle(styleTHICKRight);// 右侧边框加粗
		cell = row8.createCell(2);
		cell.setCellStyle(styleDataSecond);
		cell.setCellValue(titles4[1]);
		// cell.setCellStyle(style1);
		for (int i = 3; i < 10; i++) {
			cell = row8.createCell(i);
			cell.setCellStyle(styleDataSecond);// 填充第二组数据样式
			cell.setCellValue(Second[2][i - 3]);// 填充第二组数据
		}
		cell = row9.createCell(0);
		cell.setCellStyle(styleTHICKRight);// 右侧边框加粗
		cell = row9.createCell(2);
		cell.setCellStyle(styleDataSecond);
		cell.setCellValue(titles4[2]);
		// cell.setCellStyle(style1);
		for (int i = 3; i < 10; i++) {
			cell = row9.createCell(i);
			cell.setCellStyle(styleDataSecond);// 填充第二组数据样式
			cell.setCellValue(Second[3][i - 3]);// 填充第二组数据
		}
		cell = row10.createCell(0);
		cell.setCellStyle(styleTHICKRight);// 右侧边框加粗
		cell = row10.createCell(2);
		cell.setCellStyle(styleDataSecond);
		cell.setCellValue(titles4[3]);
		// cell.setCellStyle(style1);
		for (int i = 3; i < 10; i++) {
			cell = row10.createCell(i);
			cell.setCellStyle(styleDataSecond);// 填充第二组数据样式
			cell.setCellValue(Second[4][i - 3]);// 填充第二组数据
		}
		// 循环设置--最后一行

		for (int i = 1; i <= 9; i++) {
			cell = row11.createCell(i);
			cell.setCellStyle(styleColorBotton);
		}
		for (int i = 3; i < 10; i++) {
			cell = row11.createCell(i);
			cell.setCellStyle(styleColorBotton);// 填充第二组数据样式
			cell.setCellValue(Second[5][i - 3]);// 填充第二组数据
		}
		cell = row11.createCell(2);
		cell.setCellStyle(styleDataSecond);
		cell.setCellValue(titles4[4]);
		cell.setCellStyle(styleColorBotton);
		cell = row11.createCell(0);
		cell.setCellStyle(styleTHICKRight);// 右侧边框加粗

		// **********************************************

		// 循环设置--次下方边线加粗

		for (int i = 1; i <= 9; i++) {
			cell = row4.createCell(i);
			cell.setCellStyle(styleTHICKBotton);
		}
		// 加粗最上方边框

		for (int i = 1; i <= 9; i++) {
			cell = row0.createCell(i);
			cell.setCellStyle(styleTHICKBotton);
		}
		// 加粗次上方边框
		// HSSFRow rowTopTwo = sheet.createRow(2);
		// for (int i = 1; i <= 9; i++) {
		// cell = rowTopTwo.createCell(i);
		// cell.setCellStyle(styleTHICKBotton);
		// }

		// **********************************************
		// 涂黑边框
		for (int i = 1; i <= 8; i++) {
			cell = row3.createCell(i);
			cell.setCellStyle(styleTHICKTop);
		}
		cell = row2.createCell(0);
		cell.setCellStyle(styleTHICKRight);
		cell = row2.createCell(9);
		cell.setCellStyle(styleTHICKRightAndBotton);
		cell = row2.createCell(7);
		cell.setCellStyle(styleDataFirst);
		cell.setCellValue(daytime[0]);
		cell = row2.createCell(8);
		cell.setCellStyle(styleDataFirst);
		cell.setCellValue(daytime[1]);
		cell = row2.createCell(1);
		cell.setCellStyle(styleDataFirst);

		cell.setCellValue(first[0]);// 填充第一条数据

		cell = row2.createCell(4);
		cell.setCellStyle(styleDataFirst);
		cell.setCellValue(first[1]);

		cell = row5.createCell(10);
		cell.setCellStyle(styleTHICKLeft);
		cell = row6.createCell(10);
		cell.setCellStyle(styleTHICKLeft);
		cell = row7.createCell(10);
		cell.setCellStyle(styleTHICKLeft);
		cell = row8.createCell(10);
		cell.setCellStyle(styleTHICKLeft);
		cell = row9.createCell(10);
		cell.setCellStyle(styleTHICKLeft);
		cell = row10.createCell(10);
		cell.setCellStyle(styleTHICKLeft);
		cell = row11.createCell(10);
		cell.setCellStyle(styleTHICKLeft);
		// 循环设置--最后一行

		for (int i = 1; i <= 9; i++) {
			cell = row12.createCell(i);
			cell.setCellStyle(styleTHICKTop);
		}
		// **************************************第二张表**************************************************


		

		HSSFSheet sheet1 = wb.createSheet(sheetString);// (daytime[2] + "交易明细").trim()
		//设置列宽
		sheet1.setColumnWidth(1, 23 * 256);
		sheet1.setColumnWidth(2, 23 * 256);
		sheet1.setColumnWidth(3, 23 * 256);
		sheet1.setColumnWidth(4, 20 * 256);
		sheet1.setColumnWidth(5, 23 * 256);
		sheet1.setColumnWidth(6, 23 * 256);
		sheet1.setColumnWidth(7, 23 * 256);
		
		sheet1.setColumnWidth(8, 20 * 256);
		sheet1.setColumnWidth(9, 20 * 256);
		sheet1.setColumnWidth(10, 20 * 256);
		sheet1.setColumnWidth(11, 20 * 256);
		sheet1.setColumnWidth(12, 20 * 256);
		sheet1.setColumnWidth(13, 20 * 256);
		sheet1.setColumnWidth(14, 20 * 256);


		

		HSSFRow rowtwo0 = sheet1.createRow(0);
		HSSFRow rowtwo1 = sheet1.createRow(1);
		HSSFRow rowtwo2 = sheet1.createRow(2);
		HSSFRow rowtwo3 = sheet1.createRow(3);
		HSSFRow rowtwo4 = sheet1.createRow(4);
		HSSFRow rowtwo5 = sheet1.createRow(5);
		
		//设置行高
		rowtwo1.setHeightInPoints(26);
	
		//.createCell(0).setCellValue("1212");
		cell = rowtwo1.createCell(1);
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue("商户名称");//商户名称
		
		cell = rowtwo1.createCell(2);
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue("交易时间");
		
		cell = rowtwo1.createCell(3);
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue("渠道流水号");

		cell = rowtwo1.createCell(4);
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue("dm订单号");

		cell = rowtwo1.createCell(5);
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue("订单总金额（元）");

		cell = rowtwo1.createCell(6);
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue("会员权益优惠（元）");

		cell = rowtwo1.createCell(7);
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue("交易手续费（元）");

		cell = rowtwo1.createCell(8);
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue("到账金额（元）");

		cell = rowtwo1.createCell(9);
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue("交易状态");

		cell = rowtwo1.createCell(10);
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue("支付渠道");

		cell = rowtwo1.createCell(11);
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue("支付场景");

		cell = rowtwo1.createCell(12);
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue("终端号");

		cell = rowtwo1.createCell(13);
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue("操作员");

		cell = rowtwo1.createCell(14);
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue("备注");		

		//**************************************放入数值
		//测试从list中取值
				AccountStatementExcel accountStatementExcel = users.get(1);
				System.out.println(accountStatementExcel.getName());
//				private String name;//商户名称
//				private String time;//交易时间
//				private String number;//渠道流水号
//				private String dmNum;//dm订单号
//				private double equipmentNum;//	订单总金额（元）
//				private double vipMoney;//会员权益优惠（元）
//				private double transactionMoney;//交易手续费（元）
//				private double accountMoney;//到账金额（元）
//				private String state;//交易状态
//				private String by;//支付渠道
//				private String scene;//支付场景
//				private String terminal;//终端号
//				private String playuser;//操作员
//				private String otherPs;//备注
				double eMoney = 0;
				double vMoney = 0;
				double tMoney = 0;
				double aMoney = 0;
				
		for(int i = 0;i < users.size();i++){
			AccountStatementExcel a = users.get(i);//直接拿这个a去点get或者set就行了
			HSSFRow rowList = sheet1.createRow(i+2);
			cell = rowList.createCell(1);//此处行数要依据实际情况
			cell.setCellStyle(styleTwoSecond);
			cell.setCellValue(a.getName());//商户名称
			
			cell = rowList.createCell(2);//此处行数要依据实际情况
			cell.setCellStyle(styleTwoSecond);
			cell.setCellValue(a.getTime());//交易时间
			
			cell = rowList.createCell(3);//此处行数要依据实际情况
			cell.setCellStyle(styleTwoSecond);
			cell.setCellValue(a.getNumber());//渠道流水号
			
			cell = rowList.createCell(4);//此处行数要依据实际情况
			cell.setCellStyle(styleTwoSecond);
			cell.setCellValue(a.getDmNum());//dm订单号
		
			cell = rowList.createCell(5);//此处行数要依据实际情况
			cell.setCellStyle(styleTwoSecond);
			cell.setCellValue(a.getEquipmentNum());//	订单总金额（元）
			eMoney=a.getEquipmentNum()+eMoney;
			
			cell = rowList.createCell(6);//此处行数要依据实际情况
			cell.setCellStyle(styleTwoSecond);
			cell.setCellValue(a.getVipMoney());//会员权益优惠（元）
			vMoney=a.getVipMoney()+vMoney;

			cell = rowList.createCell(7);//此处行数要依据实际情况
			cell.setCellStyle(styleTwoSecond);
			cell.setCellValue(a.getTransactionMoney());//交易手续费（元）
			tMoney=a.getTransactionMoney()+tMoney;
					
			cell = rowList.createCell(8);//此处行数要依据实际情况
			cell.setCellStyle(styleTwoSecond);
			cell.setCellValue(a.getAccountMoney());//到账金额（元）
			aMoney=a.getAccountMoney()+aMoney;
					
			cell = rowList.createCell(9);//此处行数要依据实际情况
			cell.setCellStyle(styleTwoSecond);
			cell.setCellValue(a.getState());//交易状态

			cell = rowList.createCell(10);//此处行数要依据实际情况
			cell.setCellStyle(styleTwoSecond);
			cell.setCellValue(a.getBy());//支付渠道

			cell = rowList.createCell(11);//此处行数要依据实际情况
			cell.setCellStyle(styleTwoSecond);
			cell.setCellValue(a.getScene());//支付场景

			cell = rowList.createCell(12);//此处行数要依据实际情况
			cell.setCellStyle(styleTwoSecond);
			cell.setCellValue(a.getTerminal());//终端号

			cell = rowList.createCell(13);//此处行数要依据实际情况
			
			cell.setCellValue(a.getPlayuser());//操作员
			cell.setCellStyle(styleTwoSecond);

			cell = rowList.createCell(14);//此处行数要依据实际情况
			cell.setCellValue(a.getOtherPs());//备注
			cell.setCellStyle(styleTwoSecond);
		}
		//共计金额
		HSSFRow Total = sheet1.createRow(users.size()+2);
		for(int i = 1;i<=14;i++){
			cell = Total.createCell(i);//此处行数要依据实际情况
			cell.setCellStyle(styleTwoFirst);
			
		}
		
		cell = Total.createCell(1);//此处行数要依据实际情况
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue("合计");
		
		cell = Total.createCell(5);//此处行数要依据实际情况
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue(eMoney);//	订单总金额（元）
		
		cell = Total.createCell(6);//此处行数要依据实际情况
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue(vMoney);//	会员权益优惠（元）
		
		cell = Total.createCell(7);//此处行数要依据实际情况
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue(tMoney);//	交易手续费（元）
		
		cell = Total.createCell(8);//此处行数要依据实际情况
		cell.setCellStyle(styleTwoFirst);
		cell.setCellValue(aMoney);//	到账金额（元）
		
		

		
		//********************************加粗边线
		// 加粗最上方边框

				for (int i = 1; i <= 14; i++) {
					cell = rowtwo0.createCell(i);
					cell.setCellStyle(styleTHICKBotton);
				}
				// 加粗最下方边框
				int BottonNum=users.size()+3;
				System.out.print(BottonNum);
				HSSFRow row = sheet1.createRow(BottonNum);
				for (int i = 1; i <= 14; i++) {
					cell = row.createCell(i);//此处行数要依据实际情况
					cell.setCellStyle(styleTHICKTop);
				}
		// for (int i = 1; i <= 2; i++) {
		// cell = sheet.createRow(i).createCell(0);
		// cell.setCellStyle(styleTHICKRight);// styleTHICKTop
		// }
		// 加粗最最左侧方边框
		// for (int i = 1; i <= 9; i++) {
		// HSSFRow rowleft = sheet.createRow(1);
		// cell = rowleft.createCell(0);
		// cell.setCellStyle(styleTHICKLeft);
		// }

		/* } */

		// HSSFSheet sheet1 = wb.createSheet();
		//
		// HSSFRow row1[] = new HSSFRow[5];
		// for (int i = 0; i < 5; i++)
		// {
		// row1[i] = sheet.createRow(i);
		// }
		//
		// HSSFCell cell1[][] = new HSSFCell[5][3];
		// for (int i = 0; i < 5; i++)
		// {
		// for (int j = 0; j < 3; j++)
		// {
		// cell1[i][j] = row1[i].createCell((short) j);
		// }
		// }
		//
		// setStyle(cell1[0][0], "DASH_DOT", HSSFCellStyle.BORDER_DASH_DOT);
		// setStyle(cell1[0][1], "DASH_DOT_DOT", HSSFCellStyle.BORDER_DASH_DOT_DOT);
		// setStyle(cell1[0][2], "DASHED", HSSFCellStyle.BORDER_DASHED);
		//
		// setStyle(cell1[1][0], "DOTTED", HSSFCellStyle.BORDER_DOTTED);
		// setStyle(cell1[1][1], "DOUBLE", HSSFCellStyle.BORDER_DOUBLE);
		// setStyle(cell1[1][2], "HAIR", HSSFCellStyle.BORDER_HAIR);
		//
		// setStyle(cell1[2][0], "MEDIUM", HSSFCellStyle.BORDER_MEDIUM);
		// setStyle(cell1[2][1], "MEDIUM_DASH_DOT",
		// HSSFCellStyle.BORDER_MEDIUM_DASH_DOT);
		// setStyle(cell1[2][2], "MEDIUM_DASH_DOT_DOT",
		// HSSFCellStyle.BORDER_MEDIUM_DASH_DOT_DOT);
		//
		// setStyle(cell1[3][0], "MEDIUM_DASHED", HSSFCellStyle.BORDER_MEDIUM_DASHED);
		// setStyle(cell1[3][1], "NONE", HSSFCellStyle.BORDER_NONE);
		// setStyle(cell1[3][2], "SLANTED_DASH_DOT",
		// HSSFCellStyle.BORDER_SLANTED_DASH_DOT);
		//
		// setStyle(cell1[4][0], "THICK", HSSFCellStyle.BORDER_THICK);
		// setStyle(cell1[4][1], "THIN", HSSFCellStyle.BORDER_THIN);
		// 设置合并单元格，可以使用Ragion或者CellRagionAddress,后者为较新版本。
		// Ragion的参数为（开始行，（short）开始列），结束行，（short）结束列）。
		// Region region0 = new Region(0, (short) 0, 2, (short) 0);
		// CellRagionAddress的参数为（开始行，结束行，开始列，结束列）。
		// 商户编号 商户名称 报表生成时间 所在时区
		CellRangeAddress cellRangeAddress1 = new CellRangeAddress(1, 1, 1, 3);
		sheet.addMergedRegion(cellRangeAddress1);
		cellRangeAddress1 = new CellRangeAddress(2, 2, 1, 3);
		sheet.addMergedRegion(cellRangeAddress1);
		cellRangeAddress1 = new CellRangeAddress(1, 1, 4, 6);
		sheet.addMergedRegion(cellRangeAddress1);
		cellRangeAddress1 = new CellRangeAddress(2, 2, 4, 6);
		sheet.addMergedRegion(cellRangeAddress1);
		cellRangeAddress1 = new CellRangeAddress(1, 1, 8, 9);
		sheet.addMergedRegion(cellRangeAddress1);
		cellRangeAddress1 = new CellRangeAddress(2, 2, 8, 9);
		sheet.addMergedRegion(cellRangeAddress1);
		cellRangeAddress1 = new CellRangeAddress(3, 4, 1, 9);
		sheet.addMergedRegion(cellRangeAddress1);
		// 全时便利店
		cellRangeAddress1 = new CellRangeAddress(5, 6, 1, 1);
		sheet.addMergedRegion(cellRangeAddress1);
		// 支付渠道
		cellRangeAddress1 = new CellRangeAddress(5, 6, 2, 2);
		sheet.addMergedRegion(cellRangeAddress1);
		// 收款类订单
		cellRangeAddress1 = new CellRangeAddress(5, 5, 3, 4);
		sheet.addMergedRegion(cellRangeAddress1);
		// 退款类订单
		cellRangeAddress1 = new CellRangeAddress(5, 5, 5, 6);
		sheet.addMergedRegion(cellRangeAddress1);
		// 结算
		cellRangeAddress1 = new CellRangeAddress(5, 5, 7, 9);
		sheet.addMergedRegion(cellRangeAddress1);
		// 时间2018-05-12
		cellRangeAddress1 = new CellRangeAddress(7, 10, 1, 1);
		sheet.addMergedRegion(cellRangeAddress1);

		return wb;
	}
//获取当前时间及时区
	public static String[] datatime() {
		Date day = new Date();

		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		SimpleDateFormat d = new SimpleDateFormat("yyyy-MM-dd");

		System.out.println(df.format(day));
		// org.apache.commons.lang3.time
		System.out.print(DateFormatUtils.format(new Date(), "z"));// ‘z’小写CST；'Z'大写 +0800

		System.out.println(DateFormatUtils.format(new Date(), "ZZ"));// 'zz'小写一样 "ZZ"大写+08:00
		String[] daytime = { df.format(day),
				DateFormatUtils.format(new Date(), "z") + DateFormatUtils.format(new Date(), "ZZ"), d.format(day) };
		return daytime;
	}
}