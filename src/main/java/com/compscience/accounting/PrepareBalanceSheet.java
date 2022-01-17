package com.compscience.accounting;

import java.io.FileOutputStream;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.LoggerFactory;

public class PrepareBalanceSheet {

	public static final org.slf4j.Logger logger = LoggerFactory.getLogger(PrepareBalanceSheet.class);
	static final String accountsFile = "C:/WORK/Techzant/2021/Technizant-2021.xlsx";
	static final String finalReport = "C:/WORK/Techzant/2021/Technizant-2021-BalanceSheet.xls";
	static XSSFWorkbook wb;
	static HSSFWorkbook reportWb;
	static Map expensesCategoryMap = new HashMap();
	static Map<String, String> projectsMap = new HashMap<String, String>();

	public static void main(String args[]) {
		try {
			new PrepareBalanceSheet().generateReport();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void generateReport() throws Exception{

		OPCPackage pkg = OPCPackage.open(accountsFile);
		wb = new XSSFWorkbook(pkg);

		buildExpensesMap();
		buildProjectsMap();

		Map expensesMap = new HashMap();

		XSSFSheet sheet = wb.getSheet("EXPENSES");
		Iterator itr = sheet.rowIterator();
		double grandExpensesTotal = 0;
		while (itr.hasNext()) {
			XSSFRow row = (XSSFRow) itr.next();

			if (row.getRowNum() == 0) {
				continue;
			}

			Iterator cellItr = row.cellIterator();
			String expenseId = null;
			double expense = 0;
			while (cellItr.hasNext()) {
				XSSFCell cell = (XSSFCell) cellItr.next();
				if (cell.getColumnIndex() == 0) {
					expenseId = cell.getStringCellValue();
				}
				if (cell.getColumnIndex() == 2) {
					expense = cell.getNumericCellValue();
				}
			}

			String category = getCategory(expenseId, false);
			if (category == null) {
				continue;
			}
			Map<String, Double> detailsMap = (Map) expensesMap.get(category);
			if (detailsMap == null) {
				detailsMap = new HashMap();
			}

			Double expenseTotal = (Double) detailsMap.get(expenseId);
			if (expenseTotal == null) {
				expenseTotal = Double.valueOf(0);
			}
			expenseTotal = Double.valueOf(expenseTotal + expense);
			detailsMap.put(expenseId, expenseTotal);
			expensesMap.put(category, detailsMap);

		}

		ArrayList<String> resultList = new ArrayList<String>();

		Iterator expenseEntries = expensesMap.entrySet().iterator();

		while (expenseEntries.hasNext()) {

			Map.Entry entry = (Map.Entry) expenseEntries.next();
			String category = (String) entry.getKey();
			logger.info("Category------------------->: " + category);
			resultList.add(category + " (" + expensesCategoryMap.get(category + "-Descrption") + ")");
			Map detailsMap = (Map) entry.getValue();
			double categoryTotal = 0;
			for (Iterator itrtr = detailsMap.keySet().iterator(); itrtr.hasNext();) {

				String expenseId = (String) itrtr.next();
				double expenseTotal = ((Double) detailsMap.get(expenseId)).doubleValue();
				logger.info(
						"ExpenseId: " + expenseId + " (" + getCategory(expenseId, true) + ") Total: " + expenseTotal);
				resultList.add(
						":" + expenseId + " (" + getCategory(expenseId, true) + "):" + roundTwoDecimals(expenseTotal));
				grandExpensesTotal += expenseTotal;
				categoryTotal += expenseTotal;
			}
			resultList.add(":::" + roundTwoDecimals(categoryTotal));
			resultList.add("");

		}
		resultList.add("");

		logger.info("Grand expenses total: " + roundTwoDecimals(grandExpensesTotal));
		resultList.add("************************:************************:Expenses Total:"
				+ roundTwoDecimals(grandExpensesTotal));
		pkg.close();

		reportWb = new HSSFWorkbook();
		createSummarySheet(resultList, "Company Expenses");

		sheet = wb.getSheet("INVOICES");
		itr = sheet.rowIterator();
		Map invoicesMap = new HashMap();
		while (itr.hasNext()) {
			XSSFRow row = (XSSFRow) itr.next();

			if (row.getRowNum() == 0) {
				continue;
			}

			Iterator cellItr = row.cellIterator();
			String projectId = null;
			double incomingAmt = 0;
			double outGoingAmt = 0;
			while (cellItr.hasNext()) {
				XSSFCell cell = (XSSFCell) cellItr.next();
				if (cell.getColumnIndex() == 1) {
					projectId = cell.getStringCellValue();
				}
				if (cell.getColumnIndex() == 3) {
					incomingAmt = cell.getNumericCellValue();
				}
				if (cell.getColumnIndex() == 6) {
					outGoingAmt = cell.getNumericCellValue();
				}
			}

			Map invoiceDeatailsMap = (Map) invoicesMap.get(projectId);
			if (invoiceDeatailsMap == null) {
				invoiceDeatailsMap = new HashMap();
			}

			Double in = (Double) invoiceDeatailsMap.get("IN");
			Double out = (Double) invoiceDeatailsMap.get("OUT");

			if (in == null) {
				in = new Double(0);
			}

			if (out == null) {
				out = new Double(0);
			}
			invoiceDeatailsMap.put("IN", new Double(in.doubleValue() + incomingAmt));
			invoiceDeatailsMap.put("OUT", new Double(out.doubleValue() + outGoingAmt));

			invoicesMap.put(projectId, invoiceDeatailsMap);
		}

		double totalIncoming = 0;
		double totalOutgoing = 0;

		resultList = new ArrayList<String>();

		for (Iterator itrtr = invoicesMap.keySet().iterator(); itrtr.hasNext();) {

			String projectId = (String) itrtr.next();
			Map detailsMap = (Map) invoicesMap.get(projectId);
			double incoming = ((Double) detailsMap.get("IN")).doubleValue();
			double outgoing = ((Double) detailsMap.get("OUT")).doubleValue();
			totalIncoming += incoming;
			totalOutgoing += outgoing;
			logger.info(projectId + " IN " + incoming + " OUT " + outgoing);

			resultList.add(projectsMap.get(projectId + "-Client") + ":" + roundTwoDecimals(incoming) + ":"
					+ projectsMap.get(projectId + "-Comp1099") + ":" + roundTwoDecimals(outgoing));

		}
		resultList.add("");
		resultList.add("Total Sales (Revenue):" + roundTwoDecimals(totalIncoming) + ":Paid 1099s:"
				+ roundTwoDecimals(totalOutgoing));

		double invoicesProfit = totalIncoming - totalOutgoing;

		resultList.add("Profit:" + roundTwoDecimals(invoicesProfit));

		logger.info("Total Incoming: " + totalIncoming + " Total Outgoing: " + totalOutgoing + " Profit: "
				+ invoicesProfit);

		createInvoicesSheet(resultList, "Invoices");

		sheet = wb.getSheet("BALANCES");
		itr = sheet.rowIterator();

		double startingBankBalance = 0;
		double endingBankBalance = 0;

		while (itr.hasNext()) {
			XSSFRow row = (XSSFRow) itr.next();

			if (row.getRowNum() == 0) {
				continue;
			}

			Iterator cellItr = row.cellIterator();
			while (cellItr.hasNext()) {
				XSSFCell cell = (XSSFCell) cellItr.next();
				if (cell.getColumnIndex() == 1 && row.getRowNum() == 1) {
					startingBankBalance = cell.getNumericCellValue();
				}
				if (cell.getColumnIndex() == 2 && row.getRowNum() == 12) {
					endingBankBalance = cell.getNumericCellValue();
				}

			}

		}

		sheet = wb.getSheet("INVESTMENT_PROFIT");
		itr = sheet.rowIterator();

		resultList = new ArrayList<String>();
		resultList.add("Bank Starting Balance: " + roundTwoDecimals(startingBankBalance));
		resultList.add("Invoices Revenue:" + roundTwoDecimals(invoicesProfit));
		resultList.add("Expenses:" + -roundTwoDecimals(grandExpensesTotal));

		double endingBalance = 0;
		endingBalance += startingBankBalance;
		endingBalance += invoicesProfit;
		endingBalance += -grandExpensesTotal;
		while (itr.hasNext()) {
			XSSFRow row = (XSSFRow) itr.next();

			if (row.getRowNum() == 0) {
				continue;
			}

			Iterator cellItr = row.cellIterator();
			double investment = 0;
			String investDescription = null;
			while (cellItr.hasNext()) {
				XSSFCell cell = (XSSFCell) cellItr.next();
				if (cell.getColumnIndex() == 3) {
					investDescription = cell.getStringCellValue();
				}
				if (cell.getColumnIndex() == 2) {
					investment = cell.getNumericCellValue();
					endingBalance += investment;
				}

			}
			resultList.add(investDescription + ":" + roundTwoDecimals(investment));
		}

		resultList.add("");
		resultList.add("");
		resultList.add("Ending Balance:" + roundTwoDecimals(endingBalance));

		if (roundTwoDecimals(endingBalance) == roundTwoDecimals(endingBankBalance)) {
			logger.info("*************** Good Job ***************");
		} else {
			logger.info("*************** Just OK ***************");
		}

		createBalanceSheet(resultList, "Balance Sheet");
		logger.info("Ending Balance: " + endingBalance);
		FileOutputStream out = new FileOutputStream(finalReport);
		reportWb.write(out);
		out.close();

	}

	public static String getCategory(String expenseId, boolean description) {
		Iterator expenseCatEntries = expensesCategoryMap.entrySet().iterator();
		while (expenseCatEntries.hasNext()) {
			Map.Entry entry = (Map.Entry) expenseCatEntries.next();
			String catId = (String) entry.getKey();
			if (catId.contains("-Descrption")) {
				continue;
			}
			Map subCatMap = (Map) entry.getValue();
			Iterator subCatEntries = subCatMap.entrySet().iterator();
			while (subCatEntries.hasNext()) {
				Map.Entry subCategory = (Map.Entry) subCatEntries.next();
				String subCatId = (String) subCategory.getKey();

				if (subCatId.equals(expenseId)) {
					if (description)
						return (String) subCategory.getValue();
					else
						return catId;
				}
			}

		}
		return null;
	}

	public static void buildExpensesMap() {

		XSSFSheet sheet = wb.getSheet("EXPENSE_CATEGORY");
		Iterator itr = sheet.rowIterator();

		while (itr.hasNext()) {
			XSSFRow row = (XSSFRow) itr.next();

			if (row.getRowNum() == 0) {
				continue;
			}

			Iterator cellItr = row.cellIterator();
			String category = null;
			String categoryDescription = null;
			while (cellItr.hasNext()) {
				XSSFCell cell = (XSSFCell) cellItr.next();
				if (cell.getColumnIndex() == 0) {
					category = cell.getStringCellValue();
				}
				if (cell.getColumnIndex() == 1) {
					categoryDescription = cell.getStringCellValue();
				}
			}

			expensesCategoryMap.put(category + "-Descrption", categoryDescription);
			Map expensesSubCategoryMap = new HashMap();
			expensesCategoryMap.put(category, expensesSubCategoryMap);
		}

		sheet = wb.getSheet("EXPENSE_SUB_CATEGORY");
		itr = sheet.rowIterator();

		while (itr.hasNext()) {
			XSSFRow row = (XSSFRow) itr.next();

			if (row.getRowNum() == 0) {
				continue;
			}

			Iterator cellItr = row.cellIterator();
			String expenseId = null;
			String categoryId = null;
			String description = null;
			while (cellItr.hasNext()) {
				XSSFCell cell = (XSSFCell) cellItr.next();
				if (cell.getColumnIndex() == 0) {
					expenseId = cell.getStringCellValue();
				}
				if (cell.getColumnIndex() == 1) {
					description = cell.getStringCellValue();
				}
				if (cell.getColumnIndex() == 2) {
					categoryId = cell.getStringCellValue();
				}
			}

			Map expensesSubCategoryMap = (Map) expensesCategoryMap.get(categoryId);
			System.out.println("categoryId: " + categoryId);
			expensesSubCategoryMap.put(expenseId, description);
			expensesCategoryMap.put(categoryId, expensesSubCategoryMap);

		}

	}

	public static void buildProjectsMap() {

		XSSFSheet sheet = wb.getSheet("PROJECTS");
		Iterator itr = sheet.rowIterator();

		while (itr.hasNext()) {
			XSSFRow row = (XSSFRow) itr.next();

			if (row.getRowNum() == 0) {
				continue;
			}

			Iterator cellItr = row.cellIterator();
			String project = null;
			String client = null;
			String comp1099 = null;
			while (cellItr.hasNext()) {
				XSSFCell cell = (XSSFCell) cellItr.next();
				if (cell.getColumnIndex() == 0) {
					project = cell.getStringCellValue();
				}
				if (cell.getColumnIndex() == 1) {
					client = cell.getStringCellValue();
				}
				if (cell.getColumnIndex() == 3) {
					comp1099 = cell.getStringCellValue();
				}
			}
			projectsMap.put(project, project);
			projectsMap.put(project + "-Client", client);
			projectsMap.put(project + "-Comp1099", comp1099);

		}

	}

	public static void createSummarySheet(ArrayList<String> al, String sheetName) throws Exception {

		HSSFSheet sheet = reportWb.createSheet(sheetName);
		HSSFRow headerRow = sheet.createRow((short) 0);
		headerRow.createCell((short) 0).setCellValue(new HSSFRichTextString("EXPENSE CATEGORY"));
		headerRow.createCell((short) 1).setCellValue(new HSSFRichTextString("EXPENSE DETAILS"));
		headerRow.createCell((short) 2).setCellValue(new HSSFRichTextString("EXPENSE AMOUNT"));
		headerRow.createCell((short) 3).setCellValue(new HSSFRichTextString("EXPENSE CATEGORY AMOUNT"));
		int rowCount = 2;
		String rec;
		for (int i = 0; i < al.size(); i++) {
			headerRow = sheet.createRow((short) rowCount);
			rowCount++;
			rec = (String) al.get(i);
			String[] colValues = rec.split(":");
			for (int j = 0; j < colValues.length; j++) {
				headerRow.createCell((short) j).setCellValue(new HSSFRichTextString(colValues[j].toUpperCase()));
			}

		}

	}

	public static void createInvoicesSheet(ArrayList<String> al, String sheetName) throws Exception {

		HSSFSheet sheet = reportWb.createSheet(sheetName);
		HSSFRow headerRow = sheet.createRow((short) 0);
		headerRow.createCell((short) 0).setCellValue(new HSSFRichTextString("INCOMING INVOICES"));
		headerRow.createCell((short) 1).setCellValue(new HSSFRichTextString("INCOMING AMOUNT"));
		headerRow.createCell((short) 2).setCellValue(new HSSFRichTextString("OUTGOING INVOICES"));
		headerRow.createCell((short) 3).setCellValue(new HSSFRichTextString("OUTGOING AMOUNT"));

		int rowCount = 2;
		String rec;
		for (int i = 0; i < al.size(); i++) {
			headerRow = sheet.createRow((short) rowCount);
			rowCount++;
			rec = (String) al.get(i);
			String[] colValues = rec.split(":");
			for (int j = 0; j < colValues.length; j++) {
				headerRow.createCell((short) j).setCellValue(new HSSFRichTextString(colValues[j].toUpperCase()));
			}

		}

	}

	public static void createBalanceSheet(ArrayList<String> al, String sheetName) throws Exception {

		HSSFSheet sheet = reportWb.createSheet(sheetName);
		HSSFRow headerRow = sheet.createRow((short) 0);
		headerRow.createCell((short) 0).setCellValue(new HSSFRichTextString("DESCRIPTION"));
		headerRow.createCell((short) 1).setCellValue(new HSSFRichTextString("AMOUNT"));

		int rowCount = 2;
		String rec;
		for (int i = 0; i < al.size(); i++) {
			headerRow = sheet.createRow((short) rowCount);
			rowCount++;
			rec = (String) al.get(i);
			String[] colValues = rec.split(":");
			for (int j = 0; j < colValues.length; j++) {
				headerRow.createCell((short) j)
						.setCellValue(new HSSFRichTextString(colValues[j].toUpperCase().toUpperCase()));
			}

		}

	}

	public static double roundTwoDecimals(double d) {
		DecimalFormat twoDForm = new DecimalFormat("#.##");
		return Double.valueOf(twoDForm.format(d));
	}

}
