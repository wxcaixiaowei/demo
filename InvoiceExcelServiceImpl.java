package com.usell.platform.web.billing;

import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.usell.platform.billing.Invoice;
import com.usell.platform.billing.InvoiceCheckRequest;
import com.usell.platform.billing.InvoiceLead;
import com.usell.platform.billing.InvoiceLeadOrderItem;
import com.usell.platform.domain.Buyer;
import com.usell.platform.domain.InvoicePeriod;
import com.usell.platform.domain.PostPayCustomerPayment;


public class InvoiceExcelServiceImpl implements InvoiceExcelService {

	private InvoiceVoBuilder invoiceVoBuilder;

	private final String[] LEADS_HEADER_ROW = new String[] {"   Email Address","Prior Fees", "Current Fees", "Prior Invoiced Amount",
			"Current Invoice Amount", "Customer Billing Inter   va   l"};

	private final String[] DEVICE_DETAILS_HEADER_ROW = new String[] {"UUID", "Email Address", "User Name",
			"Invoice Period", "Order Date", "Product Name", "Proasdlkjlkasdduct Category", "Product Condition", "Device Fee", "Partner Product Id", "Partner Name"};

	private final String[] RESENT_PACK_HEADER_ROW = new String[] {"UUID","Email", "Customer Name", "Ship Date",
			"Order Date", "Product Name", "Category Name"};

	private final String[] SENT_PACK_HEADER_ROW = new String[] {"UUID","Email", "Customer Name", "Ship Date",
			"Order Date", "Product Name", "Category Name"};

	private final String[] CHECK_PROCESSING_DETAILS_HEADER_ROW = new String[] {"UUID", "Check Number", "Check Date", "Customer Name", "Amount"};

	private final String[] POST_PAY_ORDERS_HEADER_ROW = new String[] {"UUID", "Email Address", "User Name", "Order Date", "Payment Date", "Product Name", "Product Category", "Product Condition", "Order Commission Percentage",	"Bid", "Offer",	"Commission Amount Due"};

	private CellStyle cellStyle;
	
	private CellStyle percentageCellStyle;

	private static final SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");

	private static final SimpleDateFormat detailsDateFormat = new SimpleDateFormat("MM/dd/yy");

	@Override
	public void exportInvoiceToExcel(Invoice invoice, InvoicePeriod invoicePeriod, Buyer buyer,
			Boolean isPowerBuyer, OutputStream out) throws Exception {

		Workbook workbook = new HSSFWorkbook();
		cellStyle = workbook.createCellStyle();
		cellStyle.setDataFormat((short) 8);
		
		percentageCellStyle = workbook.createCellStyle();
		percentageCellStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##%"));

		addInvoiceSummary(workbook, invoiceVoBuilder.buildInvoiceSummaryVo(invoice, invoicePeriod, buyer));
		
		if (invoice.getLeads() != null && !invoice.getLeads().isEmpty()) {
			addInvoiceLeads(workbook, invoiceVoBuilder.buildLeadVo(invoice.getLeads()));	
		}
		
		List<InvoiceLeadOrderItem> invoiceOrderItems = new ArrayList<InvoiceLeadOrderItem>();
		for (InvoiceLead invoiceLead : invoice.getLeads()) {
			invoiceOrderItems.addAll(invoiceLead.getOrderItems());
		}
		
		if (!invoiceOrderItems.isEmpty()) {
			addDeviceDetails(workbook, invoiceVoBuilder.buildInvoiceLeadOrderItemVo(invoiceOrderItems, buyer.getName()), isPowerBuyer);	
		}
		
		if (invoice.getPostPayCustomerPayments() != null && !invoice.getPostPayCustomerPayments().isEmpty()) {
			addPostPaidOrders(workbook, invoice.getPostPayCustomerPayments());	
		}
		
		List<InvoiceKitVo> sentKitsVo = invoiceVoBuilder.buildSentKitsVo(invoice.getShippingKits());
		if (sentKitsVo != null && !sentKitsVo.isEmpty()) {
			addSentKitDetails(workbook, sentKitsVo);	
		}
		
		List<InvoiceKitVo> resentKitsVo = invoiceVoBuilder.buildReshippedVo(invoice.getShippingKits());
		if (resentKitsVo != null && !resentKitsVo.isEmpty()) {
			addReshipKitDetails(workbook, resentKitsVo);	
		}

		if (invoice.getCheckRequests() != null && !invoice.getCheckRequests().isEmpty()) {
			addInvoiceCheckRequests(workbook, invoice.getCheckRequests());	
		}
		
		workbook.write(out);
		out.flush();
		out.close();
	}

	private void addPostPaidOrders(Workbook workBook,
			List<PostPayCustomerPayment> postPayCustomerPayments) {

		Sheet postPayOrderSheet = workBook.createSheet("Orders");
		int rowIndex = 0;
		Row topRow = postPayOrderSheet.createRow(rowIndex++);
		populateHeaderRow(getStandardHeaderStyle(workBook), POST_PAY_ORDERS_HEADER_ROW, topRow);
		
		rowIndex = 1;
		for (PostPayCustomerPayment payment : postPayCustomerPayments) {
			int colIndex = 0;
			Row dataRow = postPayOrderSheet.createRow(rowIndex);
			Cell dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(payment.getOrderNumber());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(payment.getEmail());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(payment.getFirstName().concat(" ".concat(payment.getLastName())).toUpperCase());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(detailsDateFormat.format(payment.getOrderDate()));

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(detailsDateFormat.format(payment.getPaymentDate()));

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(payment.getProductName());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(payment.getProductCategoryName());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(payment.getProductConditionName());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(payment.getOrderCommissionPercentage()+"%");

			dataCell = createNewCell(dataRow, colIndex++, cellStyle);
			dataCell.setCellValue(payment.getFinalBid());
			dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);

			dataCell = createNewCell(dataRow, colIndex++, cellStyle);
			dataCell.setCellValue(payment.getFinalOffer());
			dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);

			dataCell = createNewCell(dataRow, colIndex++, cellStyle);
			dataCell.setCellValue(payment.getOrderCommission());
			dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);
			rowIndex++;
		}

		//AutoSize columns
		for (int i=0; i < topRow.getLastCellNum(); i++){
			postPayOrderSheet.autoSizeColumn(i);
		}
	}

	private void addInvoiceCheckRequests(Workbook workBook,
			List<InvoiceCheckRequest> checkRequests) {

		if(checkRequests.size() < 1) {
			return;
		}

		Sheet checkRequestSheet = workBook.createSheet("Check Processed");
		int rowIndex = 0;
		Row topRow = checkRequestSheet.createRow(rowIndex++);
		populateHeaderRow(getStandardHeaderStyle(workBook), CHECK_PROCESSING_DETAILS_HEADER_ROW, topRow);

		rowIndex = 1;
		for (InvoiceCheckRequest request : checkRequests) {
			int colIndex = 0;
			Row dataRow = checkRequestSheet.createRow(rowIndex);

			Cell dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(request.getOrderUid());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(request.getCheckNumber());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(dateFormat.format(request.getCheckDate()));

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(request.getFirstName().concat(" ").concat(request.getLastName()));

			dataCell = createNewCell(dataRow, colIndex++, cellStyle);
			dataCell.setCellValue(request.getCheckAmount());
			dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);
			rowIndex++;
		}
		//AutoSize columns
		for (int i=0; i < topRow.getLastCellNum(); i++){
			checkRequestSheet.autoSizeColumn(i);
		}

	}

	private void addInvoiceSummary(Workbook workbook, InvoiceSummaryVo invoiceSummary){
		Sheet summarySheet = workbook.createSheet("Invoice Summary");
		int rowIndex = 0;

		Row topRow = summarySheet.createRow(rowIndex++);

		int colIndex = 0;
		Cell dataCell = createNewCell(topRow, colIndex++, null);
		dataCell.setCellValue("Partner Name");

		dataCell = createNewCell(topRow, colIndex++, null);
		dataCell.setCellValue(invoiceSummary.getBuyer());
		colIndex++;

		dataCell = createNewCell(topRow, colIndex++, null);
		dataCell.setCellValue("Invoice:");

		dataCell = createNewCell(topRow, colIndex++, null);
		dataCell.setCellValue(invoiceSummary.getInvoiceNumber());

		Row dataRow = summarySheet.createRow(rowIndex++);

		colIndex = 0;
		dataCell = createNewCell(dataRow, colIndex++, null);
		dataCell.setCellValue("Invoice Start Date");

		dataCell = createNewCell(dataRow, colIndex++, null);
		dataCell.setCellValue(invoiceSummary.getInvoicePeriodStartDate());

		dataRow = summarySheet.createRow(rowIndex++);

		colIndex = 0;
		dataCell = createNewCell(dataRow, colIndex++, null);
		dataCell.setCellValue("Invoice End Date");

		dataCell = createNewCell(dataRow, colIndex++, null);
		dataCell.setCellValue(invoiceSummary.getInvoicePeriodEndDate());

		dataRow = summarySheet.createRow(rowIndex++);

		colIndex = 0;
		dataCell = createNewCell(dataRow, colIndex++, null);
		dataCell.setCellValue("Date Generated");

		dataCell = createNewCell(dataRow, colIndex++, null);
		dataCell.setCellValue(invoiceSummary.getInvoiceDate());

		rowIndex++;
		dataRow = summarySheet.createRow(rowIndex++);
		colIndex = 0;
		dataCell = createNewCell(dataRow, colIndex++, null);
		dataCell.setCellValue("Charges for Post Pay Orders");

		dataRow = summarySheet.createRow(rowIndex++);
		colIndex = 1;
		dataCell = createNewCell(dataRow, colIndex++, null);
		dataCell.setCellValue("Number of Orders");

		dataCell = createNewCell(dataRow, colIndex++, null);
		dataCell.setCellValue(invoiceSummary.getPostPayOrderCount());

		dataRow = summarySheet.createRow(rowIndex++);
		colIndex = 1;
		dataCell = createNewCell(dataRow, colIndex++, null);
		dataCell.setCellValue("Net Amount Due");

		dataCell = createNewCell(dataRow, colIndex++, cellStyle);
		dataCell.setCellValue(invoiceSummary.getTotalPostPayOrderCommission());
		dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);


		if (invoiceSummary.getShippingKitSummaryVo() != null) {
			rowIndex++;
			dataRow = summarySheet.createRow(rowIndex++);
			colIndex = 0;
			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue("Charges for Shipping Kits");

			dataRow = summarySheet.createRow(rowIndex++);
			colIndex = 1;
			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue("Date");

			for (String category : invoiceSummary.getShippingKitSummaryVo().getProductCategories()) {
				dataCell = createNewCell(dataRow, colIndex++, null);
				dataCell.setCellValue(category);
			}

			for (String day : invoiceSummary.getShippingKitSummaryVo().getDays()) {
				dataRow = summarySheet.createRow(rowIndex++);
				colIndex = 1;
				dataCell = createNewCell(dataRow, colIndex++, null);
				dataCell.setCellValue(day);

				for (String category : invoiceSummary.getShippingKitSummaryVo().getProductCategories()) {
					ShippingKitSummaryVo.CategoryDay categoryDay = new ShippingKitSummaryVo.CategoryDay();
					categoryDay.day = day;
					categoryDay.productCategory = category;
					dataCell = createNewCell(dataRow, colIndex++, null);
					dataCell.setCellValue(invoiceSummary.getShippingKitSummaryVo().getDailyCategoryCount().get(categoryDay));
				}
			}

			rowIndex++;
			dataRow = summarySheet.createRow(rowIndex++);
			colIndex = 1;
			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue("Count");

			for (String category : invoiceSummary.getShippingKitSummaryVo().getProductCategories()) {
				dataCell = createNewCell(dataRow, colIndex++, null);
				dataCell.setCellValue(invoiceSummary.getShippingKitSummaryVo().getCategoryCount().get(category));
			}

			dataRow = summarySheet.createRow(rowIndex++);
			colIndex = 1;
			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue("Price Per Unit");

			for (String category : invoiceSummary.getShippingKitSummaryVo().getProductCategories()) {
				dataCell = createNewCell(dataRow, colIndex++, cellStyle);
				dataCell.setCellValue(Double.parseDouble(invoiceSummary.getShippingKitSummaryVo().getCategoryPricePerUnit().get(category)));
				dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);

			}

			dataRow = summarySheet.createRow(rowIndex++);
			colIndex = 1;
			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue("Amount Due");

			for (String category : invoiceSummary.getShippingKitSummaryVo().getProductCategories()) {
				dataCell = createNewCell(dataRow, colIndex++, cellStyle);
				dataCell.setCellValue(Double.parseDouble(invoiceSummary.getShippingKitSummaryVo().getCategoryAmountDue().get(category)));
				dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);

			}

			rowIndex++;
			dataRow = summarySheet.createRow(rowIndex++);
			colIndex = 1;
			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue("Net Amount Due");

			dataCell = createNewCell(dataRow, colIndex++, cellStyle);
			dataCell.setCellValue(Double.parseDouble(invoiceSummary.getShippingKitSummaryVo().getAmountDue()));
			dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);

		}

		if(invoiceSummary.getTotalCheckCount() > 0) {
			rowIndex++;
			dataRow = summarySheet.createRow(rowIndex++);
			colIndex = 0;
			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue("Charges for Check Processing");

			dataRow = summarySheet.createRow(rowIndex++);
			colIndex = 1;
			dataCell = createNewCell(dataRow, colIndex, null);
			dataCell.setCellValue("Count");
			dataCell = createNewCell(dataRow, colIndex + 1, null);
			dataCell.setCellValue(invoiceSummary.getTotalCheckCount());

			dataRow = summarySheet.createRow(rowIndex++);
			dataCell = createNewCell(dataRow, colIndex, null);
			dataCell.setCellValue("Price Per Unit");
			dataCell = createNewCell(dataRow, colIndex + 1, cellStyle);
			dataCell.setCellValue(invoiceSummary.getCheckChargePerUnit());
			dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);

			dataRow = summarySheet.createRow(rowIndex++);
			dataCell = createNewCell(dataRow, colIndex, null);
			dataCell.setCellValue("Amount Due");
			dataCell = createNewCell(dataRow, colIndex + 1, cellStyle);
			dataCell.setCellValue(invoiceSummary.getTotalCheckAmountDue());
			dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);
		}

		rowIndex++;
		dataRow = summarySheet.createRow(rowIndex++);
		colIndex = 0;
		dataCell = createNewCell(dataRow, colIndex++, null);
		dataCell.setCellValue("TOTAL AMOUNT DUE");

		dataCell = createNewCell(dataRow, colIndex++, cellStyle);
		dataCell.setCellValue(Double.parseDouble(invoiceSummary.getAmountDue()));
		dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);

		for (int i=0; i < topRow.getLastCellNum(); i++){
			summarySheet.autoSizeColumn(i);
		}

	}


	private void addDeviceDetails(Workbook workbook, List<InvoiceLeadOrderItemVo> deviceDetailsVos, Boolean isBuyer) {
		Sheet deviceLevelDetails = workbook.createSheet("Prepaid Device Level Details");
		Row headerRow = deviceLevelDetails.createRow(0);
		populateDeviceDetailsHeaderRow(getStandardHeaderStyle(workbook), DEVICE_DETAILS_HEADER_ROW, headerRow, isBuyer);

		int rowIndex = 1;
		for(InvoiceLeadOrderItemVo deviceDetails : deviceDetailsVos){
			int colIndex = 0;
			Row dataRow = deviceLevelDetails.createRow(rowIndex);

			Cell dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(deviceDetails.getUuid());

			dataCell = createNewCell(dataRow, colIndex++, cellStyle);
			dataCell.setCellValue(deviceDetails.getCustomerEmail());

			dataCell = createNewCell(dataRow, colIndex++, cellStyle);
			dataCell.setCellValue(deviceDetails.getCustomerFullName());

			dataCell = createNewCell(dataRow, colIndex++, cellStyle);
			dataCell.setCellValue(deviceDetails.getInvoicePeriod());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(deviceDetails.getOrderDate());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(deviceDetails.getProductName());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(deviceDetails.getProductCategory());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(deviceDetails.getProductCondition());

			dataCell = createNewCell(dataRow, colIndex++, cellStyle);
			dataCell.setCellValue(deviceDetails.getDeviceFee());

			if(!isBuyer) {
				dataCell = createNewCell(dataRow, colIndex++, null);
				dataCell.setCellValue(deviceDetails.getPartnerProductId());
			}

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(deviceDetails.getPartnerName());

			rowIndex++;
		}

		//AutoSize columns
		for (int i=0; i < headerRow.getLastCellNum(); i++){
			deviceLevelDetails.autoSizeColumn(i);
		}
	}

	private void addInvoiceLeads(Workbook workbook, List<InvoiceLeadVo> invoiceLeads) {
		Sheet leadsSheet = workbook.createSheet("Leads");
		Row headerRow = leadsSheet.createRow(0);
		populateHeaderRow(getStandardHeaderStyle(workbook), LEADS_HEADER_ROW, headerRow);
		int rowIndex = 1;
		for(InvoiceLeadVo lead : invoiceLeads){
			int colIndex = 0;
			Row dataRow = leadsSheet.createRow(rowIndex);

			Cell dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(lead.getCustomerEmail());

			dataCell = createNewCell(dataRow, colIndex++, cellStyle);
			dataCell.setCellValue(lead.getPriorCommission());
			dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);

			dataCell = createNewCell(dataRow, colIndex++, cellStyle);
			dataCell.setCellValue(lead.getCurrentCommission());
			dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);

			dataCell = createNewCell(dataRow, colIndex++, cellStyle);
			dataCell.setCellValue(lead.getPriorFees());
			dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);

			dataCell = createNewCell(dataRow, colIndex++, cellStyle);
			dataCell.setCellValue(lead.getCurrentFees());
			dataCell.setCellType(Cell.CELL_TYPE_NUMERIC);

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(lead.getBillingInterval());
			rowIndex++;
		}

		//AutoSize columns
		for (int i=0; i < headerRow.getLastCellNum(); i++){
			leadsSheet.autoSizeColumn(i);
		}
	}

	private void addSentKitDetails(Workbook workbook,	List<InvoiceKitVo> invoiceReshippedKits) {
		Sheet reshipSheet = workbook.createSheet("Kits Sent");
		Row headerRow = reshipSheet.createRow(0);
		populateHeaderRow(getStandardHeaderStyle(workbook), SENT_PACK_HEADER_ROW, headerRow);
		int rowIndex = 1;
		for(InvoiceKitVo invoiceReshippedKit : invoiceReshippedKits){
			int colIndex = 0;
			Row dataRow = reshipSheet.createRow(rowIndex);

			Cell dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(invoiceReshippedKit.getOrderNumber());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(invoiceReshippedKit.getCustomerEmail());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(invoiceReshippedKit.getCustomerName());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(invoiceReshippedKit.getShipDate());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(invoiceReshippedKit.getOrderDate());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(invoiceReshippedKit.getProductName());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(invoiceReshippedKit.getProductCategoryName());
			rowIndex++;
		}

		//AutoSize columns
		for (int i=0; i < headerRow.getLastCellNum(); i++){
			reshipSheet.autoSizeColumn(i);
		}
	}

	
	private void addReshipKitDetails(Workbook workbook,	List<InvoiceKitVo> invoiceReshippedKits) {
		Sheet reshipSheet = workbook.createSheet("Kits Resent");
		Row headerRow = reshipSheet.createRow(0);
		populateHeaderRow(getStandardHeaderStyle(workbook), RESENT_PACK_HEADER_ROW, headerRow);
		int rowIndex = 1;
		for(InvoiceKitVo invoiceReshippedKit : invoiceReshippedKits){
			int colIndex = 0;
			Row dataRow = reshipSheet.createRow(rowIndex);

			Cell dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(invoiceReshippedKit.getOrderNumber());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(invoiceReshippedKit.getCustomerEmail());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(invoiceReshippedKit.getCustomerName());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(invoiceReshippedKit.getShipDate());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(invoiceReshippedKit.getOrderDate());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(invoiceReshippedKit.getProductName());

			dataCell = createNewCell(dataRow, colIndex++, null);
			dataCell.setCellValue(invoiceReshippedKit.getProductCategoryName());
			rowIndex++;
		}

		//AutoSize columns
		for (int i=0; i < headerRow.getLastCellNum(); i++){
			reshipSheet.autoSizeColumn(i);
		}
	}

	private Cell createNewCell(Row row, int index, CellStyle cellStyle){
		Cell cell = row.createCell(index);
		if(cellStyle != null){
			cell.setCellStyle(cellStyle);
		}
		return cell;
	}

	private void populateHeaderRow(CellStyle style, String[] colHeaders, Row row){
		for(int i=0; i<colHeaders.length; i++ ){
			Cell headerCell = createNewCell(row, i, style);
			headerCell.setCellValue(colHeaders[i]);
		}
	}

	private void populateDeviceDetailsHeaderRow(CellStyle style, String[] deviceDetailsHeader, Row row, Boolean isBuyer){

		for(int i=0,j=0; i<deviceDetailsHeader.length; i++,j++ ){

			if(isBuyer && deviceDetailsHeader[i].equalsIgnoreCase("Partner Product Id")) {
				j--;
				continue;
			}

			Cell headerCell = createNewCell(row, j, style);
			headerCell.setCellValue(deviceDetailsHeader[i]);
		}
	}

	private CellStyle getStandardHeaderStyle(Workbook workBook){
		Font font = workBook.createFont();
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		CellStyle cellStyle = workBook.createCellStyle();
		cellStyle.setFont(font);
		cellStyle.setAlignment(HSSFCellStyle.ALIGN_LEFT);
		return cellStyle;
	}
	public void setInvoiceVoBuilder(InvoiceVoBuilder invoiceVoBuilder) {
		this.invoiceVoBuilder = invoiceVoBuilder;
	}
}
