package com.compscience.accounting;

import java.io.File;
import java.io.FileInputStream;
import java.text.DecimalFormat;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.LoggerFactory;

public class AccountsValidator {

	public static final org.slf4j.Logger logger = LoggerFactory.getLogger(AccountsValidator.class);
	static final String ACCOUNTS_FILE = "G:/Other computers/valis-pc/work/Techzant/2023/business/Technizant-2023.xlsx";

	public static void main(String[] args) {

		new AccountsValidator().validator();
	}

	public void validator() {
		
		logger.info("Validation started...");
		
		try {
			FileInputStream excelFile = new FileInputStream(new File(ACCOUNTS_FILE));
			XSSFWorkbook wb = new XSSFWorkbook(excelFile);
			XSSFSheet sheet = wb.getSheet("EXPENSES");
			Iterator<Row> itr = sheet.rowIterator();
			double totalExpenses = 0;
			while (itr.hasNext()) {
				XSSFRow row = (XSSFRow) itr.next();
				if (row.getRowNum() == 0) {
					continue;
				}
				Iterator<Cell> cellItr = row.cellIterator();
				while (cellItr.hasNext()) {
					XSSFCell cell = (XSSFCell) cellItr.next();
					if (cell.getColumnIndex() == 2) {
						totalExpenses += cell.getNumericCellValue();
					}
				}
			}
			logger.info("Expenses total: " + totalExpenses);

			sheet = wb.getSheet("BALANCES");
			itr = sheet.rowIterator();
			double beginBal = 0;
			double endBal = 0;
			while (itr.hasNext()) {
				XSSFRow row = (XSSFRow) itr.next();
				if (row.getRowNum() == 0) {
					continue;
				}
				Iterator<Cell> cellItr = row.cellIterator();
				while (cellItr.hasNext()) {
					XSSFCell cell = (XSSFCell) cellItr.next();
					if (cell.getColumnIndex() == 1) {
						beginBal = cell.getNumericCellValue();
					}
					if (cell.getColumnIndex() == 2) {
						endBal = cell.getNumericCellValue();
					}
				}
			}

			logger.info("Begin Balance: " + beginBal);
			logger.info("End Balance: " + endBal);

			sheet = wb.getSheet("INVOICES");
			itr = sheet.rowIterator();
			double invoicesBalance = 0;
			while (itr.hasNext()) {
				XSSFRow row = (XSSFRow) itr.next();
				if (row.getRowNum() == 0) {
					continue;
				}
				Iterator<Cell> cellItr = row.cellIterator();
				while (cellItr.hasNext()) {
					XSSFCell cell = (XSSFCell) cellItr.next();
					if (cell.getColumnIndex() == 3) {
						invoicesBalance += cell.getNumericCellValue();
					}
					if (cell.getColumnIndex() == 6) {
						invoicesBalance -= cell.getNumericCellValue();
					}
				}
			}

			logger.info("Invoices Balance: " + invoicesBalance);

			sheet = wb.getSheet("INVESTMENT_PROFIT");
			itr = sheet.rowIterator();
			double investments = 0;
			while (itr.hasNext()) {
				XSSFRow row = (XSSFRow) itr.next();
				if (row.getRowNum() == 0) {
					continue;
				}
				Iterator<Cell> cellItr = row.cellIterator();
				while (cellItr.hasNext()) {
					XSSFCell cell = (XSSFCell) cellItr.next();
					if (cell.getColumnIndex() == 2) {
						investments += cell.getNumericCellValue();
					}
				}
			}

			logger.info("Investments: " + investments);

			double profit = investments + invoicesBalance + beginBal;
			double calculatedEndingBalance = roundTwoDecimals(profit - totalExpenses);
			double difference = endBal - calculatedEndingBalance;
			if (difference < 1.0 && difference >= 0) {
				logger.info("************* Hurray. You are the accounting champion! *************");
				logger.info("Calculated Ending Balance: "+calculatedEndingBalance);
				logger.info("End Balance: "+endBal);
				logger.info("Difference: " + difference);
				logger.info("********************************************************************");
			} else {
				logger.info("Sorry. Something is wrong with your accounts :(");
				logger.info("Profit minus expenses: " + calculatedEndingBalance);
				logger.info("Ending balance: " + endBal);
				logger.info("Difference: " + difference);
			}

			wb.close();
			//pkg.close();

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	double roundTwoDecimals(double d) {
		DecimalFormat twoDForm = new DecimalFormat("#.##");
		return Double.valueOf(twoDForm.format(d));
	}

}
