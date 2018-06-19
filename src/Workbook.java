import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Workbook {
	public static final String TEMPLATE_XLSX_FILE_PATH = "/temp/PlanilhaTemplate.xlsx";
	public static final String PLANILHA_GERADA_XLSX_FILE_PATH = "/temp/PlanilhaGerada.xlsx";
	public static final Integer COLUNA_INICIO_DADOS_RELATORIO = 1;
	public static final Integer COLUNA_FIM_DADOS_RELATORIO = 5;
	public static final Integer LINHA_INICIO_DADOS_RELATORIO = 4;
	public static Double soma = 0.0d;
	public static Integer qdRegistro = 0;

	public static void main(String[] args) throws EncryptedDocumentException, InvalidFormatException, IOException {
		readAndWriteXLSXFile();
	}

	public static void readAndWriteXLSXFile() throws EncryptedDocumentException, InvalidFormatException, IOException {

		montaPlanilhaRelatorio();
	}

	private static void montaPlanilhaRelatorio()
			throws EncryptedDocumentException, InvalidFormatException, IOException {
		InputStream inp = new FileInputStream(TEMPLATE_XLSX_FILE_PATH);
		XSSFWorkbook wb = (XSSFWorkbook) WorkbookFactory.create(inp);
		// Aba Template - começa em 0
		XSSFSheet sheetTemplate = wb.getSheetAt(0);
		// Aba Dados
		XSSFSheet sheetDados = wb.getSheetAt(1);
		// Aba Dados
		XSSFSheet sheetRelatorio = wb.createSheet("Relatório Company Poipoi");

		montaCabecalho(sheetRelatorio, sheetTemplate);
		montaCorpoRelatorio(sheetRelatorio, sheetTemplate, sheetDados);
		montaRodape(sheetRelatorio, sheetDados);

		// Formatação para célula Data que é montada no cabeçalho
		CreationHelper createHelper = wb.getCreationHelper();
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yyyy"));
		sheetRelatorio.getRow(0).getCell(7).setCellStyle(cellStyle);

		// Abrir o Excel na planilha do template com os valores populados
		sheetRelatorio.setSelected(true);

		escrevePlanilha(wb);
	}

	private static void escrevePlanilha(XSSFWorkbook wb) throws FileNotFoundException {

		FileOutputStream fileOut = new FileOutputStream(TEMPLATE_XLSX_FILE_PATH);

		System.out.println("writing...");
		// write this workbook to an Outputstream.
		try {
			wb.write(fileOut);
			fileOut.flush();
			fileOut.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static XSSFSheet montaCabecalho(XSSFSheet sheetRelatorio, XSSFSheet sheetTemplate) {
		// Data
		Row linhaUm = sheetRelatorio.createRow(0);
		Cell celulaData = linhaUm.createCell(7);
		celulaData.setCellValue(new Date());

		// Título relatório
		Row linhaDois = sheetRelatorio.createRow(1);
		Cell celulaMescladaTitulo = linhaDois.createCell(1);
		celulaMescladaTitulo.setCellValue("Relatório Caterpillar Teste");

		// Mesclando segunda linha para Título
		sheetRelatorio.addMergedRegion(new CellRangeAddress(1, 1, 1, 5));// 1°
																			// linha
																			// até
																			// 1°
																			// linha,
																			// 1°coluna
																			// até
																			// 5°coluna

		// Linha do Cabecalho
		Row linhaCabecalhoTemplate = sheetTemplate.getRow(0);
		Row linhaCabecalhoRelatorio = sheetRelatorio.createRow(2);
		for (Cell cellTemplate : linhaCabecalhoTemplate) {
			Cell cellRelatorio = linhaCabecalhoRelatorio.createCell(cellTemplate.getColumnIndex());
			cellRelatorio.setCellValue(cellTemplate.getStringCellValue());
		}

		return sheetRelatorio;
	}

	private static XSSFSheet montaCorpoRelatorio(XSSFSheet sheetRelatorio, XSSFSheet sheetTemplate,
			XSSFSheet sheetDados) {
		setClienteRelatorio(sheetRelatorio, sheetTemplate, sheetDados);
		Integer linhaFinalRelatorio = totalRegistros(sheetDados) + 4;
		insereFiltroRelatorio(sheetRelatorio, LINHA_INICIO_DADOS_RELATORIO, linhaFinalRelatorio,
				COLUNA_INICIO_DADOS_RELATORIO, COLUNA_FIM_DADOS_RELATORIO);

		return sheetRelatorio;
	}

	private static XSSFSheet setClienteRelatorio(XSSFSheet sheetRelatorio, XSSFSheet sheetTemplate,
			XSSFSheet sheetDados) {
		int i = 0;
		for (Row rowDados : sheetDados) {

			if (i > sheetDados.getPhysicalNumberOfRows())
				break;

			if (sheetDados.getLastRowNum() > 0)
				i++;

			// linha relatório
			Row rowRel = sheetRelatorio.createRow(i + 3);

			for (Cell cellDados : rowDados) {

				Cell cellRelatorio = rowRel.createCell(cellDados.getColumnIndex() + 1);
				insereValorCellPorTipo(cellRelatorio, cellDados);
			}
		}
		return sheetRelatorio;
	}

	private static void insereValorCellPorTipo(Cell cellRelatorio, Cell cellDados) {

		switch (cellDados.getCellTypeEnum()) {

		case BOOLEAN:
			cellRelatorio.setCellValue(cellDados.getBooleanCellValue());
			break;
		case STRING:
			cellRelatorio.setCellValue(cellDados.getRichStringCellValue().getString());
			break;
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cellDados)) {
				cellRelatorio.setCellValue(cellDados.getDateCellValue());
			} else {
				cellRelatorio.setCellValue(cellDados.getNumericCellValue());
			}
			break;
		case FORMULA:
			cellRelatorio.setCellValue(cellDados.getCellFormula());
			break;
		case BLANK:
			System.out.print("");
			break;
		default:
			System.out.print("");
		}

	}

	private static XSSFSheet montaRodape(XSSFSheet sheetRelatorio, XSSFSheet sheetDados) {
		// Somatório das Operações na primeira Aba
		XSSFSheet sheet = setValorSomatorio(sheetRelatorio, sheetDados);
		return sheet;

	}

	private static String copiaFormulaCell(XSSFSheet sheet, int row, int column) {
		return sheet.getRow(row).getCell(column).getCellFormula();
	}

	private static XSSFSheet setValorSomatorio(XSSFSheet sheetRelatorio, XSSFSheet sheetDados) {
		Integer totalLinhasRegistro = totalRegistros(sheetDados);
		Cell cellTituloTotal = sheetRelatorio.createRow(totalLinhasRegistro + 10).createCell(0);
		cellTituloTotal.setCellValue("Total ");
		Cell cellValorTotal = sheetRelatorio.getRow(totalLinhasRegistro + 10).createCell(1);
		cellValorTotal.setCellValue(totalOperacoes(sheetDados));
		return sheetRelatorio;
	}

	private static Double totalOperacoes(XSSFSheet sheetDados) {
		sheetDados.forEach(row -> {
			soma += row.getCell(4).getNumericCellValue();
		});

		return soma;
	}

	private static Integer totalRegistros(XSSFSheet sheet) {
		return sheet.getPhysicalNumberOfRows();
	}

	private static XSSFSheet colaFormulaCell(XSSFSheet sheetOrigemFormula, int rowOri, int colOri,
			XSSFSheet sheetDestinoFormula, int rowDest, int colDest) {
		String formulaSoma = copiaFormulaCell(sheetOrigemFormula, rowOri, colOri);
		sheetDestinoFormula.getRow(rowDest).getCell(colDest).setCellFormula(formulaSoma);
		return sheetDestinoFormula;
	}

	private static XSSFSheet insereFiltroRelatorio(XSSFSheet sheetRelatorio, int linhaInicial, int linhaFinal,
			int colInicial, int colFinal) {
		sheetRelatorio.setAutoFilter(new CellRangeAddress(linhaInicial, linhaFinal, colInicial, colFinal));

		return sheetRelatorio;

	}

	private static XSSFSheet formataCellsCurrency(XSSFSheet sheet, int rowInicio, int rowFinal, int colInicio,
			int colFinal) {
		return sheet;
	}
}
