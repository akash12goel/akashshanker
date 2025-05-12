import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import java.time.Year;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Class to store grouped data for each Business Group and Environment Name
 * combination
 */
class GroupData {
	double flows;
	double messages;
	double throughput;
	double workers;
	String businessGroup;
	String environmentName;

	public GroupData(String businessGroup, String environmentName) {
		this.flows = 0;
		this.messages = 0;
		this.throughput = 0;
		this.workers = 0;
		this.businessGroup = businessGroup;
		this.environmentName = environmentName;
	}

	/**
	 * Adds numeric values to the running totals, skipping invalid or negative
	 * values
	 */
	public void addValues(double flows, double messages, double throughput, double workers) {
		if (flows >= 0)
			this.flows += flows;
		if (messages >= 0)
			this.messages += messages;
		if (throughput >= 0)
			this.throughput += throughput;
		if (workers >= 0)
			this.workers += workers;
	}

	public Object[] toRowData() {
		return new Object[] { businessGroup, environmentName, flows, messages, throughput, workers };
	}
}

/**
 * Class to store grouped data at Business Group level
 */
class BusinessGroupSummary {
	@Override
	public String toString() {
		return "BusinessGroupSummary [flows=" + flows + ", messages=" + messages + ", throughput=" + throughput
				+ ", workers=" + workers + ", apiCount=" + apiCount + ", cost=" + cost + ", businessGroup="
				+ businessGroup + "]";
	}

	double flows;
	double messages;
	double throughput;
	double workers;
	double apiCount;
	double cost;
	String businessGroup;

	public BusinessGroupSummary(String businessGroup) {
		this.flows = 0;
		this.messages = 0;
		this.throughput = 0;
		this.workers = 0;
		this.apiCount = 0;
		this.cost = 0;
		this.businessGroup = businessGroup;
	}

	public void addValues(double flows, double messages, double throughput, double workers) {
		if (flows >= 0)
			this.flows += flows;
		if (messages >= 0)
			this.messages += messages;
		if (throughput >= 0)
			this.throughput += throughput;
		if (workers >= 0)
			this.workers += workers;
	}

	public void setApiCount(double apiCount) {
		if (apiCount >= 0)
			this.apiCount = apiCount;
	}

	public void calculateCost(double flowCost, double messageCost, double throughputCost, double apiCost) {
		this.cost = (this.flows * flowCost) + (this.messages * messageCost) + (this.throughput * throughputCost)
				+ (this.apiCount * apiCost);
	}

	public Object[] toRowData() {
		return new Object[] { businessGroup, flows, messages, throughput, workers, apiCount, cost };
	}
}

/**
 * Main class to process CSV data and group by Business Group and Environment
 * Name
 */
public class CSVProcessor {
	private static final String MULE_FLOWS = "Mule Flows";
	// MONTHS will be determined dynamically in main()
	private static final String[] HEADERS = { "Business Group", "Environment Name", MULE_FLOWS, "Mule Messages",
			"Data Throughput", "Worker Count" };
	private static final String[] SUMMARY_HEADERS = { "Business Group", "Owner Name", "Track Lead", MULE_FLOWS,
			"Mule Messages", "Data Throughput(GB) ", "Worker Count", "# of APIs Managed", "Cost($)" };

	// For Total Cost Summary
	private static final String[] TOTAL_COST_HEADERS = { "Business Group", "Owner Name", "Track Lead", "Cost($)" };

	// Helper class for ownership details
	static class OwnershipDetails {
		String ownerName;
		String trackLead;

		public OwnershipDetails(String ownerName, String trackLead) {
			this.ownerName = ownerName;
			this.trackLead = trackLead;
		}
	}

	/**
	 * MuleSoft Spend Report Generator
	 *
	 * This program processes monthly MuleSoft runtime and API manager CSV files,
	 * aggregates costs and usage by business group and track lead, and generates
	 * a multi-sheet Excel report with summaries, totals, and charts.
	 *
	 * Key features:
	 * - Dynamically detects months from input files
	 * - Produces month-wise, business group, and track lead summaries
	 * - Adds grand total rows and bold formatting for totals
	 * - Uses Calibri font for all data
	 * - Generates a line chart for cost variation
	 */
	public static void main(String[] args) {
		// Set the report year (can be dynamic with Year.now().getValue())
		int year = Year.now().getValue();
		String baseInputPath = "C:\\Users\\ashanker2\\Documents\\workspace-spring-tool-suite-4-4.21.0.RELEASE\\JavaExp\\src\\";
		String monthlyExcelPath = baseInputPath + "Mule_Runtime_Monthly_Report_" + year + ".xlsx";
		String summaryExcelPath = baseInputPath + "MuleSoft Spend by Product & Owners " + year + ".xlsx";
		String muleCostPath = baseInputPath + "Mule_cost.csv";
		String ownershipPath = baseInputPath + "Mule_Ownership_Details.csv";

		// Read cost values
		Map<String, Double> costs = new HashMap<>();
		try {
			costs = readCostFile(muleCostPath);
			costs.forEach((key, value) -> {
				System.out.println("Key: " + key + ", Value: " + value);
			});
		} catch (IOException e) {
			System.err.println("Error reading cost file: " + e.getMessage());
			e.printStackTrace();
			return;
		}

		// Read ownership details
		Map<String, OwnershipDetails> ownershipMap = new HashMap<>();
		try {
			ownershipMap = readOwnershipFile(ownershipPath);

		} catch (IOException e) {
			System.err.println("Error reading ownership file: " + e.getMessage());
			e.printStackTrace();
			return;
		}

		// Dynamically determine months from files in the input directory
		File dir = new File(baseInputPath);
		String prefix = "Mule_Runtime_Monthly_";
		String suffix = ".csv";
		List<String> monthsList = new ArrayList<>();
		for (File file : dir.listFiles()) {
			String name = file.getName();
			if (name.startsWith(prefix) && name.endsWith(suffix)) {
				String month = name.substring(prefix.length(), name.length() - suffix.length());
				monthsList.add(month);
			}
		}
		// Sort months in reverse calendar order (Dec, Nov, ..., Jan)
		List<String> calendarOrder = Arrays.asList("January", "February", "March", "April", "May", "June", "July",
				"August", "September", "October", "November", "December");
		monthsList.sort((m1, m2) -> {
			int idx1 = calendarOrder.indexOf(m1);
			int idx2 = calendarOrder.indexOf(m2);
			if (idx1 == -1)
				idx1 = Integer.MAX_VALUE;
			if (idx2 == -1)
				idx2 = Integer.MAX_VALUE;
			return Integer.compare(idx2, idx1); // reverse order
		});
		String[] MONTHS = monthsList.toArray(new String[0]);

		// Process data for both files
		Map<String, Map<String, Map<String, GroupData>>> allMonthData = new HashMap<>();
		Map<String, Map<String, BusinessGroupSummary>> monthlyBusinessGroupSummary = new HashMap<>();

		// First, collect all data
		for (String month : MONTHS) {
			String inputCsvPath = baseInputPath + "Mule_Runtime_Monthly_" + month + ".csv";
			String apiManagerPath = baseInputPath + "API_Manager_Monthly_" + month + ".csv";
			try {
				// Process Mule Runtime data
				Map<String, Map<String, GroupData>> monthData = processMonthData(inputCsvPath);
				allMonthData.put(month, monthData);

				// Create separate business group summary for each month
				Map<String, BusinessGroupSummary> monthSummary = new TreeMap<>(); // Using TreeMap for automatic sorting
				updateBusinessGroupSummary(monthData, monthSummary);

				// Add API count data
				updateApiCount(apiManagerPath, monthSummary);

				// Calculate costs for each business group
				for (BusinessGroupSummary summary : monthSummary.values()) {
					summary.calculateCost(costs.getOrDefault(MULE_FLOWS.toString(), 0.0),
							costs.getOrDefault("Mule Messages", 0.0), costs.getOrDefault("Data Throughput", 0.0),
							costs.getOrDefault("# of APIs Managed", 0.0));
				}

				monthlyBusinessGroupSummary.put(month, monthSummary);
				monthlyBusinessGroupSummary.forEach((key, value) -> {
					System.out.println("Key: " + key + ", Value: " + value);
				});
			} catch (IOException e) {
				System.err.println("Error processing " + month + " data: " + e.getMessage());
				e.printStackTrace();
				return;
			}
		}

		// Create monthly report
		try (Workbook monthlyWorkbook = new XSSFWorkbook()) {
			for (String month : MONTHS) {
				Sheet monthSheet = monthlyWorkbook.createSheet(month);
				createHeader(monthlyWorkbook, monthSheet, HEADERS);
				writeDataToSheet(monthlyWorkbook, monthSheet, allMonthData.get(month));
			}

			// Auto-size columns
			for (int i = 0; i < monthlyWorkbook.getNumberOfSheets(); i++) {
				Sheet sheet = monthlyWorkbook.getSheetAt(i);
				for (int col = 0; col < HEADERS.length; col++) {
					sheet.autoSizeColumn(col);
				}
			}

			// Write monthly workbook to file
			try (FileOutputStream fileOut = new FileOutputStream(monthlyExcelPath)) {
				monthlyWorkbook.write(fileOut);
			}
			System.out.println("Monthly Excel file created successfully at: " + monthlyExcelPath);

		} catch (IOException e) {
			System.err.println("Error creating monthly report: " + e.getMessage());
			e.printStackTrace();
		}

		// Create business group summary report with separate sheets for each month
		try (Workbook summaryWorkbook = new XSSFWorkbook()) {
			// Create a Calibri font style for all data cells
			Font calibriFont = summaryWorkbook.createFont();
			calibriFont.setFontName("Calibri");
			CellStyle calibriStyle = summaryWorkbook.createCellStyle();
			calibriStyle.setFont(calibriFont);

			// Create a bold Calibri style for Grand Total rows
			Font boldCalibriFont = summaryWorkbook.createFont();
			boldCalibriFont.setFontName("Calibri");
			boldCalibriFont.setBold(true);
			CellStyle boldCalibriStyle = summaryWorkbook.createCellStyle();
			boldCalibriStyle.setFont(boldCalibriFont);

			// Create a sheet for each month (month-wise summary)
			for (String month : MONTHS) {
				Sheet monthSheet = summaryWorkbook.createSheet(month);
				createHeader(summaryWorkbook, monthSheet, SUMMARY_HEADERS);
				writeSummaryToSheetWithOwnership(summaryWorkbook, monthSheet, monthlyBusinessGroupSummary.get(month),
						ownershipMap, calibriStyle);

				// Auto-size columns
				for (int col = 0; col < SUMMARY_HEADERS.length; col++) {
					monthSheet.autoSizeColumn(col);
				}
			}

			// --- Add Total Cost Summary sheet ---
			// Aggregate all months into a business group summary
			Map<String, BusinessGroupSummary> totalSummary = new TreeMap<>();
			for (Map<String, BusinessGroupSummary> monthSummary : monthlyBusinessGroupSummary.values()) {
				for (Map.Entry<String, BusinessGroupSummary> entry : monthSummary.entrySet()) {
					String businessGroup = entry.getKey();
					BusinessGroupSummary monthData = entry.getValue();
					BusinessGroupSummary total = totalSummary.computeIfAbsent(businessGroup,
							k -> new BusinessGroupSummary(businessGroup));
					total.flows += monthData.flows;
					total.messages += monthData.messages;
					total.throughput += monthData.throughput;
					total.workers += monthData.workers;
					total.apiCount += monthData.apiCount;
					total.cost += monthData.cost;
				}
			}
			// Create the sheet and write Business Group, Owner, Lead, month-wise costs, and Total Cost
			Sheet totalSheet = summaryWorkbook.createSheet("Total Cost Summary");
			// Prepare headers: Business Group, Owner Name, Track Lead, [Month1], [Month2],
			// ...
			String[] monthCostHeaders = new String[4 + MONTHS.length];
			monthCostHeaders[0] = "Business Group";
			monthCostHeaders[1] = "Owner Name";
			monthCostHeaders[2] = "Track Lead";
			for (int i = 0; i < MONTHS.length; i++) {
				monthCostHeaders[3 + i] = MONTHS[i] + " Cost($)";
			}
			monthCostHeaders[3 + MONTHS.length] = "Total Cost($)";
			createHeader(summaryWorkbook, totalSheet, monthCostHeaders);
			int rowNum = 1;
			for (BusinessGroupSummary summary : totalSummary.values()) {
				Row row = totalSheet.createRow(rowNum++);
				OwnershipDetails od = ownershipMap.getOrDefault(summary.businessGroup.trim().toLowerCase(),
						new OwnershipDetails("", ""));
				Cell cell0 = row.createCell(0);
				cell0.setCellValue(summary.businessGroup);
				cell0.setCellStyle(calibriStyle);
				Cell cell1 = row.createCell(1);
				cell1.setCellValue(od.ownerName);
				cell1.setCellStyle(calibriStyle);
				Cell cell2 = row.createCell(2);
				cell2.setCellValue(od.trackLead);
				cell2.setCellStyle(calibriStyle);
				// Add month-wise costs
				for (int i = 0; i < MONTHS.length; i++) {
					Map<String, BusinessGroupSummary> monthSummary = monthlyBusinessGroupSummary.get(MONTHS[i]);
					double monthCost = 0.0;
					if (monthSummary != null && monthSummary.containsKey(summary.businessGroup)) {
						monthCost = monthSummary.get(summary.businessGroup).cost;
					}
					Cell cell = row.createCell(3 + i);
					cell.setCellValue(Math.round(monthCost * 100.0) / 100.0);
					cell.setCellStyle(calibriStyle);
				}
				Cell grandTotalRow = row.createCell(3 + MONTHS.length);
				grandTotalRow.setCellValue(Math.round(summary.cost * 100.0) / 100.0);
				grandTotalRow.setCellStyle(boldCalibriStyle);
			}
			// Add a total row for the sum of all the month cost columns and total cost
			Row grandTotalRow = totalSheet.createRow(rowNum++);
			Cell totalLabelCell = grandTotalRow.createCell(0);
			totalLabelCell.setCellValue("Grand Total");
			totalLabelCell.setCellStyle(boldCalibriStyle);
			// Empty cells for Owner Name and Track Lead
			for (int i = 1; i < 3; i++) {
				Cell cell = grandTotalRow.createCell(i);
				cell.setCellValue("");
				cell.setCellStyle(boldCalibriStyle);
			}
			// Sum each month cost column and total cost column
			for (int i = 0; i < MONTHS.length + 1; i++) {
				int colIdx = 3 + i;
				double sum = 0.0;
				for (int r = 1; r < rowNum - 1; r++) { // skip header, exclude total row
					Row row = totalSheet.getRow(r);
					if (row != null) {
						Cell cell = row.getCell(colIdx);
						if (cell != null) {
							sum += cell.getNumericCellValue();
						}
					}
				}
				Cell sumCell = grandTotalRow.createCell(colIdx);
				sumCell.setCellValue(Math.round(sum * 100.0) / 100.0);
				sumCell.setCellStyle(boldCalibriStyle);
			}
			for (int col = 0; col < monthCostHeaders.length; col++) {
				totalSheet.autoSizeColumn(col);
			}

			// --- Create a new sheet for the Cost Variation Chart ---
			XSSFSheet chartSheet = (XSSFSheet) summaryWorkbook.createSheet("Cost Variation Chart");
			// Only chart, no data table. Use data from totalSheet.
			int lastRow = rowNum - 1;
			int firstDataRow = 1;
			int lastDataRow = lastRow;
			XSSFDrawing drawing = chartSheet.createDrawingPatriarch();
			ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 1, 10, 20);
			XSSFChart chart = drawing.createChart(anchor);
			chart.setTitleText("Cost Variation per Business Group (per Month)");
			chart.setTitleOverlay(false);
			XDDFChartLegend legend = chart.getOrAddLegend();
			legend.setPosition(LegendPosition.TOP_RIGHT);
			XDDFCategoryAxis xAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
			xAxis.setTitle("Business Group");
			XDDFValueAxis yAxis = chart.createValueAxis(AxisPosition.LEFT);
			yAxis.setTitle("Cost");
			// Use data from totalSheet: Business Group in col 0, months in cols
			// 3..(3+MONTHS.length-1)
			CellRangeAddress catRange = new CellRangeAddress(firstDataRow, lastDataRow, 0, 0);
			XDDFCategoryDataSource categories = XDDFDataSourcesFactory.fromStringCellRange((XSSFSheet) totalSheet,
					catRange);
			XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, xAxis, yAxis);
			for (int m = 0; m < MONTHS.length; m++) {
				CellRangeAddress valRange = new CellRangeAddress(firstDataRow, lastDataRow, 4 + m, 4 + m);
				XDDFNumericalDataSource<Double> values = XDDFDataSourcesFactory
						.fromNumericCellRange((XSSFSheet) totalSheet, valRange);
				XDDFLineChartData.Series series = (XDDFLineChartData.Series) data.addSeries(categories, values);
				series.setTitle(MONTHS[m], null);
				series.setSmooth(false);
				series.setMarkerStyle(MarkerStyle.CIRCLE);
			}
			chart.plot(data);
			// --- End Cost Variation Chart Sheet ---

			// --- Add Track Lead Spend Summary sheet ---
			// Map to accumulate total spend per track lead
			Map<String, Double> trackLeadSpend = new HashMap<>();
			for (BusinessGroupSummary summary : totalSummary.values()) {
				OwnershipDetails od = ownershipMap.getOrDefault(summary.businessGroup.trim().toLowerCase(),
						new OwnershipDetails("", ""));
				String trackLead = od.trackLead;
				double spend = Math.round(summary.cost * 100.0) / 100.0;
				trackLeadSpend.put(trackLead, trackLeadSpend.getOrDefault(trackLead, 0.0) + spend);
			}
			Sheet leadSheet = summaryWorkbook.createSheet("Track Lead Spend Summary");
			String[] leadHeaders = { "Track Lead", "Budget $(" + year + ")", "Total Spend $(YTD)" };
			createHeader(summaryWorkbook, leadSheet, leadHeaders);
			int leadRowNum = 1;
			double grandTotal = 0.0;
			for (Map.Entry<String, Double> entry : trackLeadSpend.entrySet()) {
				Row row = leadSheet.createRow(leadRowNum++);
				Cell cell0 = row.createCell(0);
				cell0.setCellValue(entry.getKey());
				cell0.setCellStyle(calibriStyle);
				double spend = Math.round(entry.getValue() * 100.0) / 100.0;
				Cell cell1 = row.createCell(1);
				cell1.setCellValue("Not Baseline Yet");
				cell1.setCellStyle(calibriStyle);
				Cell cell2 = row.createCell(2);
				cell2.setCellValue(spend);
				cell2.setCellStyle(calibriStyle);
				grandTotal += spend;
			}
			// Add grand total row (bold) and optional notes
			Row totalRow = leadSheet.createRow(leadRowNum++);
			Cell leadTotalLabelCell = totalRow.createCell(0);
			leadTotalLabelCell.setCellValue("Grand Total");
			leadTotalLabelCell.setCellStyle(boldCalibriStyle);
			Cell leadTotalBudgetCell = totalRow.createCell(1);
			leadTotalBudgetCell.setCellValue("300,000");
			leadTotalBudgetCell.setCellStyle(boldCalibriStyle);
			Cell leadTotalSpendCell = totalRow.createCell(2);
			leadTotalSpendCell.setCellValue(Math.round(grandTotal * 100.0) / 100.0);
			leadTotalSpendCell.setCellStyle(boldCalibriStyle);
			Row spareRow1 = leadSheet.createRow(leadRowNum++);
			Row spareRow2 = leadSheet.createRow(leadRowNum++);
			Row note = leadSheet.createRow(leadRowNum++);
			//Cell noteCell0 = note.createCell(0);
			//noteCell0.setCellValue("Note");
			//noteCell0.setCellStyle(calibriStyle);
			//Cell noteCell1 = note.createCell(1);
		//	noteCell1.setCellValue("In addition to these expenses, there are $120K for the platform and $30K for the VPC.");
			//noteCell1.setCellStyle(calibriStyle);

			for (int col = 0; col < leadHeaders.length; col++) {
				leadSheet.autoSizeColumn(col);
			}

			// Reorder sheets: Total Cost Summary, Cost Variation Chart, Track Lead Spend
			// Summary, then month-wise sheets
			summaryWorkbook.setSheetOrder("Track Lead Spend Summary", 0);
			summaryWorkbook.setSheetOrder("Total Cost Summary", 1);
			summaryWorkbook.setSheetOrder("Cost Variation Chart", 2);
			summaryWorkbook.setFirstVisibleTab(0);
			// Write summary workbook to file
			try (FileOutputStream fileOut = new FileOutputStream(summaryExcelPath)) {
				summaryWorkbook.write(fileOut);
			}
			System.out.println("Business Group Summary Excel file created successfully at: " + summaryExcelPath);

		} catch (IOException e) {
			System.err.println("Error creating business group summary: " + e.getMessage());
			e.printStackTrace();
		}
	}

	/**
	 * Creates a header row with orange background and bold Calibri font.
	 */
	private static void createHeader(Workbook workbook, Sheet sheet, String[] headers) {
		Row headerRow = sheet.createRow(0);
		CellStyle headerStyle = workbook.createCellStyle();
		headerStyle.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
		headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		Font headerFont = workbook.createFont();
		headerFont.setBold(true);
		headerFont.setFontName("Calibri");
		headerStyle.setFont(headerFont);

		for (int i = 0; i < headers.length; i++) {
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(headers[i]);
			cell.setCellStyle(headerStyle);
		}
	}

	/**
	 * Aggregates environment-level data into business group summaries for a month.
	 */
	private static void updateBusinessGroupSummary(Map<String, Map<String, GroupData>> monthData,
			Map<String, BusinessGroupSummary> businessGroupSummary) {

		for (Map.Entry<String, Map<String, GroupData>> entry : monthData.entrySet()) {
			String businessGroup = entry.getKey();
			BusinessGroupSummary summary = businessGroupSummary.computeIfAbsent(businessGroup,
					k -> new BusinessGroupSummary(businessGroup));

			for (GroupData data : entry.getValue().values()) {
				summary.addValues(data.flows, data.messages, data.throughput, data.workers);
			}
		}
	}

	/**
	 * Writes a month-wise summary sheet with business group, owner, lead, and cost columns.
	 */
	private static void writeSummaryToSheetWithOwnership(Workbook workbook, Sheet sheet,
			Map<String, BusinessGroupSummary> businessGroupSummary, Map<String, OwnershipDetails> ownershipMap, CellStyle calibriStyle) {
		int rowNum = 1; // Start after header
		for (BusinessGroupSummary summary : businessGroupSummary.values()) {
			Row row = sheet.createRow(rowNum++);
			OwnershipDetails od = ownershipMap.getOrDefault(summary.businessGroup.trim().toLowerCase(),
					new OwnershipDetails("", ""));
			Cell cell0 = row.createCell(0);
			cell0.setCellValue(summary.businessGroup);
			cell0.setCellStyle(calibriStyle);
			Cell cell1 = row.createCell(1);
			cell1.setCellValue(od.ownerName);
			cell1.setCellStyle(calibriStyle);
			Cell cell2 = row.createCell(2);
			cell2.setCellValue(od.trackLead);
			cell2.setCellStyle(calibriStyle);
			Cell cell3 = row.createCell(3);
			cell3.setCellValue(summary.flows);
			cell3.setCellStyle(calibriStyle);
			Cell cell4 = row.createCell(4);
			cell4.setCellValue(summary.messages);
			cell4.setCellStyle(calibriStyle);
			Cell cell5 = row.createCell(5);
			cell5.setCellValue(summary.throughput);
			cell5.setCellStyle(calibriStyle);
			Cell cell6 = row.createCell(6);
			cell6.setCellValue(summary.workers);
			cell6.setCellStyle(calibriStyle);
			Cell cell7 = row.createCell(7);
			cell7.setCellValue(summary.apiCount);
			cell7.setCellStyle(calibriStyle);
			Cell cell8 = row.createCell(8);
			cell8.setCellValue(Math.round(summary.cost * 100.0) / 100.0);
			cell8.setCellStyle(calibriStyle);
		}
	}

	/**
	 * Reads and groups CSV data for a month by business group and environment.
	 */
	private static Map<String, Map<String, GroupData>> processMonthData(String inputCsvPath) throws IOException {
		Map<String, Map<String, GroupData>> businessGroupData = new TreeMap<>();

		try (BufferedReader br = new BufferedReader(new FileReader(inputCsvPath))) {
			String line;
			boolean isFirstLine = true;
			int[] columnIndices = new int[6];

			while ((line = br.readLine()) != null) {
				String[] values = line.split(",");

				if (isFirstLine) {
					for (int i = 0; i < values.length; i++) {
						String header = values[i].trim();
						header = header.replaceAll("[\\p{Cf}]", "");
						switch (header) {
						case "Business Group":
							columnIndices[0] = i;
							break;
						case "Environment Name":
							columnIndices[1] = i;
							break;
						case MULE_FLOWS:
							columnIndices[2] = i;
							break;
						case "Mule Messages":
							columnIndices[3] = i;
							break;
						case "Data Throughput":
							columnIndices[4] = i;
							break;
						case "Worker Count":
							columnIndices[5] = i;
							break;
						}
					}
					isFirstLine = false;
					continue;
				}

				processRow(values, columnIndices, businessGroupData);
			}
		}

		return businessGroupData;
	}

	/**
	 * Processes a single row of CSV data and updates the group data map.
	 */
	private static void processRow(String[] values, int[] indices,
			Map<String, Map<String, GroupData>> businessGroupData) {
		String businessGroup = values[indices[0]].trim();
		String environmentName = values[indices[1]].trim();

		double flows = parseDoubleOrZero(values[indices[2]]);
		double messages = parseDoubleOrZero(values[indices[3]]);
		double throughput = parseDoubleOrZero(values[indices[4]]);
		double workers = parseDoubleOrZero(values[indices[5]]);

		Map<String, GroupData> envMap = businessGroupData.computeIfAbsent(businessGroup, k -> new TreeMap<>());
		GroupData groupData = envMap.computeIfAbsent(environmentName,
				k -> new GroupData(businessGroup, environmentName));
		groupData.addValues(flows, messages, throughput, workers);
	}

	/**
	 * Writes grouped environment data to a sheet for a given month.
	 */
	private static void writeDataToSheet(Workbook workbook, Sheet sheet,
			Map<String, Map<String, GroupData>> businessGroupData) {
		int rowNum = 1; // Start after header

		for (Map<String, GroupData> envMap : businessGroupData.values()) {
			for (GroupData data : envMap.values()) {
				Row row = sheet.createRow(rowNum++);
				Object[] rowData = data.toRowData();

				for (int i = 0; i < rowData.length; i++) {
					Cell cell = row.createCell(i);
					if (i <= 1) { // Business Group and Environment Name are strings
						cell.setCellValue((String) rowData[i]);
					} else { // Numeric values
						cell.setCellValue((Double) rowData[i]);
					}
				}
			}
		}
	}

	/**
	 * Parses a string value to double, handling invalid values.
	 * Returns -1 for invalid or N/A values to indicate they should be skipped.
	 */
	private static double parseDoubleOrZero(String value) {
		if (value == null || value.trim().isEmpty() || value.trim().equals("N/A")) {
			return -1; // Return -1 to indicate invalid/missing values
		}
		try {
			double parsed = Double.parseDouble(value.trim());
			return parsed >= 0 ? parsed : -1; // Only accept non-negative values
		} catch (NumberFormatException e) {
			return -1;
		}
	}

	/**
	 * Updates the business group summary with API counts from the API manager CSV.
	 */
	private static void updateApiCount(String apiManagerPath, Map<String, BusinessGroupSummary> businessGroupSummary)
			throws IOException {
		// First, collect and group API counts by Business Group
		Map<String, Double> groupedApiCounts = new TreeMap<>(); // Using TreeMap for automatic sorting

		try (BufferedReader br = new BufferedReader(new FileReader(apiManagerPath))) {
			String line;
			boolean isFirstLine = true;
			int businessGroupIndex = -1;
			int apiCountIndex = -1;

			while ((line = br.readLine()) != null) {
				String[] values = line.split(",");

				if (isFirstLine) {
					// Find column indices
					for (int i = 0; i < values.length; i++) {
						String header = values[i].trim();
						header = header.replaceAll("[\\p{Cf}]", "");
						if (header.equals("Business Group")) {
							businessGroupIndex = i;
						} else if (header.equals("# of APIs Managed")) {
							apiCountIndex = i;
						}
					}
					isFirstLine = false;
					continue;
				}

				if (businessGroupIndex >= 0 && apiCountIndex >= 0) {
					String businessGroup = values[businessGroupIndex].trim();
					double apiCount = parseDoubleOrZero(values[apiCountIndex]);

					if (apiCount >= 0) {
						// Add to the grouped counts
						groupedApiCounts.merge(businessGroup, apiCount, Double::sum);
					}
				}
			}
		}

		// Now update the business group summary with the grouped API counts
		for (Map.Entry<String, Double> entry : groupedApiCounts.entrySet()) {
			String businessGroup = entry.getKey();
			double totalApiCount = entry.getValue();

			// Get or create the BusinessGroupSummary
			BusinessGroupSummary summary = businessGroupSummary.computeIfAbsent(businessGroup,
					k -> new BusinessGroupSummary(businessGroup));

			// Set the API count
			summary.setApiCount(totalApiCount);
		}
	}

	/**
	 * Reads the cost file and returns a map of metric to cost value.
	 */
	private static Map<String, Double> readCostFile(String costFilePath) throws IOException {
		Map<String, Double> costs = new HashMap<>();
		try (BufferedReader br = new BufferedReader(new FileReader(costFilePath))) {
			String headerLine = br.readLine();
			String valueLine = br.readLine();
			if (headerLine == null || valueLine == null) {
				return costs;
			}
			String[] headers = headerLine.split(",");
			String[] values = valueLine.split(",");
			for (int i = 0; i < headers.length && i < values.length; i++) {
				String metric = headers[i].trim();
				metric = metric.replaceAll("[\\p{Cf}]", "");
				double cost = parseDoubleOrZero(values[i]);
				if (cost >= 0) {
					costs.put(metric.toString(), cost);
				}
			}
		}
		return costs;
	}

	/**
	 * Reads the ownership details file and returns a map of business group to owner/lead.
	 */
	private static Map<String, OwnershipDetails> readOwnershipFile(String ownershipPath) throws IOException {
		Map<String, OwnershipDetails> map = new HashMap<>();
		try (BufferedReader br = new BufferedReader(new FileReader(ownershipPath))) {
			String line;
			boolean isFirstLine = true;
			int businessGroupIdx = -1, ownerIdx = -1, leadIdx = -1;
			while ((line = br.readLine()) != null) {
				String[] values = line.split(",");
				if (isFirstLine) {
					for (int i = 0; i < values.length; i++) {
						String header = values[i].trim();
						header = header.replaceAll("[\\p{Cf}]", "");
						if (header.equalsIgnoreCase("Business Group"))
							businessGroupIdx = i;
						else if (header.equalsIgnoreCase("Owner Name"))
							ownerIdx = i;
						else if (header.equalsIgnoreCase("Track Lead"))
							leadIdx = i;
					}
					isFirstLine = false;
					continue;
				}
				// businessGroupIdx = 0;
				if (businessGroupIdx >= 0 && ownerIdx >= 0 && leadIdx >= 0) {
					String bg = values[businessGroupIdx].trim().toLowerCase();
					String owner = values[ownerIdx].trim();
					String lead = values[leadIdx].trim();
					map.put(bg, new OwnershipDetails(owner, lead));
				}
			}
		}
		return map;
	}
}