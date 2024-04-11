package com.puerto.bobinas.informes.helpers;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.LinkOption;
import java.nio.file.Path;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Calendar;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.stream.Collectors;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFTextbox;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Component;

import com.puerto.bobinas.informes.beans.Bobina;
import com.puerto.bobinas.informes.beans.BobinasTemplate;
import com.puerto.bobinas.informes.constantes.Constantes;
import com.puerto.bobinas.informes.enums.ClientesEnum;
import com.puerto.bobinas.informes.utils.Utilidades;

import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class ExcelHelper {

	private static final String VAR_CAB = "VAR_CAB";

	private static final String VAR_PESO_BRUTO = "VAR_PESO_BRUTO";

	private static final String VAR_NUM_SERIE = "VAR_NUM_SERIE";

	private static final String VAR_POSITION = "VAR_POSITION";

	private static final String VAR_DESTINATARIO = "VAR_DESTINATARIO";

	private static final String VAR_VESSEL = "VAR_VESSEL";

	private static final String VAR_FECHA = "VAR_FECHA";

	private static final int ARCELOR_BLOQUE_HEADER_SIZE = 13;

	private static final int ARCELOR_BLOQUE_FOOTER_SIZE = 15;

	private static final int ARCELOR_BLOQUE_MAXIMOS = 43;

	private static final int MARGIN_PANELES_TEXT = 20160;

	@Autowired
	private Utilidades utilidades;

	@Value("${fileChooser.directory.root}")
	private String rootDirectory;
	@Value("${fileChooser.directory.excels.entrada}")
	private String entradaDirectory;
	@Value("${fileChooser.directory.excels.salida}")
	private String salidaDirectory;
	@Value("${fileChooser.dialog.title}")
	private String fileChooserDialogTitle;
	@Value("${tableView.bobinas.serie.alias}")
	private String serieAlias;
	@Value("${tableView.bobinas.destinatario.alias}")
	private String destinatarioAlias;
	@Value("${tableView.bobinas.pesoBruto.alias}")
	private String pesoBrutoAlias;
	@Value("classpath:plantillas/plantilla_thyssen.xlsx")
	private Resource plantillaThyssen;
	@Value("classpath:plantillas/plantilla_arcelor.xlsx")
	private Resource plantillaArcelor;
	@Value("classpath:img/firma.jpg")
	private Resource imgFirma;

	public FileChooser getFileChooserEntrada() {
		final FileChooser fileChooser = new FileChooser();
		ExtensionFilter ex1 = new ExtensionFilter("Excel File", "*.xlsx", "*.xls");
		ExtensionFilter ex2 = new ExtensionFilter("all Files", "*.*");
		fileChooser.getExtensionFilters().addAll(ex1, ex2);
		fileChooser.setTitle(fileChooserDialogTitle);
		if (utilidades.crearDirectorio(entradaDirectory)) {
			log.info("Directorio {} creado", entradaDirectory);
		}
		File entradaDirectoryFile = FileUtils.getFile(entradaDirectory);
		if (entradaDirectoryFile != null && FileUtils.isDirectory(entradaDirectoryFile, LinkOption.NOFOLLOW_LINKS)) {
			fileChooser.setInitialDirectory(entradaDirectoryFile);
		}
		return fileChooser;
	}

	public BobinasTemplate getBobinasTemplate(String pathString) {
		var bobinasTemplate = new BobinasTemplate();
		var bobinas = new ArrayList<Bobina>();
		try {
			var serieColPos = -1;
			var destinatarioColPos = -1;
			var pesoBrutoColPos = -1;
			File f = new File(pathString);
			InputStream is = new FileInputStream(f);
			Workbook wb = WorkbookFactory.create(is);
			Sheet sheet = wb.getSheetAt(0);
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				DecimalFormat decimalFormat = new DecimalFormat(Constantes.PATTERN_NUMBER_STRING);
				switch (row.getRowNum()) {
				case Constantes.EXCEL_ROW_POS_ENCABEZADO:
					try {
						Iterator<Cell> cellIteratorEncabezado = row.cellIterator();
						var encabezadoString = new StringBuilder();
						while (cellIteratorEncabezado.hasNext()) {
							Cell cellEncabezado = cellIteratorEncabezado.next();
							if (cellEncabezado.getCellType().equals(CellType.STRING)) {
								encabezadoString.append(cellEncabezado.getStringCellValue());
							} else if (cellEncabezado.getCellType().equals(CellType.NUMERIC)) {
								encabezadoString.append(decimalFormat.format(cellEncabezado.getNumericCellValue()));
							} else {
								encabezadoString.append(StringUtils.EMPTY);
							}
						}
						bobinasTemplate.setEncabezado(encabezadoString.toString());
						bobinasTemplate.setBarco(obtenerBarco(bobinasTemplate));
						var clientesStringList = Arrays.asList(ClientesEnum.values()).stream()
								.map(ClientesEnum::getValor).collect(Collectors.toList());
						clientesStringList.forEach(clienteString -> {
							if (bobinasTemplate.getEncabezado().contains(clienteString)) {
								bobinasTemplate.setCliente(clienteString);
							}
						});
					} catch (IllegalStateException e) {
						bobinasTemplate.setCliente(StringUtils.EMPTY);
					}
					break;
				case Constantes.EXCEL_ROW_POS_CABECERAS:
					Iterator<Cell> cellIterator = row.cellIterator();
					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();
						try {
							String stringCellValue = cell.getStringCellValue();
							if (StringUtils.containsIgnoreCase(stringCellValue, serieAlias)) {
								serieColPos = cell.getColumnIndex();
							}
							if (StringUtils.containsIgnoreCase(stringCellValue, destinatarioAlias)) {
								destinatarioColPos = cell.getColumnIndex();
							}
							if (StringUtils.containsIgnoreCase(stringCellValue, pesoBrutoAlias)) {
								pesoBrutoColPos = cell.getColumnIndex();
							}
						} catch (IllegalStateException e) {
							log.error("Cabecera con valor numerico");
						}
					}
					break;
				default:
					var bobina = new Bobina();
					Cell destinatarioCell = row.getCell(destinatarioColPos);
					Cell serieCell = row.getCell(serieColPos);
					Cell pesoBrutoCell = row.getCell(pesoBrutoColPos);
					var rowBlank = destinatarioCell == null & serieCell == null;
					if (!rowBlank) {
						if (destinatarioCell.getCellType().equals(CellType.STRING)) {
							bobina.setNombreDestinatario(destinatarioCell.getStringCellValue());
						} else if (destinatarioCell.getCellType().equals(CellType.NUMERIC)) {
							bobina.setNombreDestinatario(decimalFormat.format(destinatarioCell.getNumericCellValue()));
						} else {
							bobina.setNombreDestinatario("Destinatario no identificado");
						}
						//
						if (serieCell.getCellType().equals(CellType.STRING)) {
							bobina.setNumSerie(serieCell.getStringCellValue());
						} else if (serieCell.getCellType().equals(CellType.NUMERIC)) {
							bobina.setNumSerie(decimalFormat.format(serieCell.getNumericCellValue()));
						} else {
							bobina.setNumSerie("Serie no identificada");
						}
						//
						if (pesoBrutoCell.getCellType().equals(CellType.STRING)
								&& NumberUtils.isDigits(pesoBrutoCell.getStringCellValue())) {
							bobina.setPesoBrutoPrevisto(Double.parseDouble((pesoBrutoCell.getStringCellValue())));
						} else if (pesoBrutoCell.getCellType().equals(CellType.NUMERIC)) {
							bobina.setPesoBrutoPrevisto(pesoBrutoCell.getNumericCellValue());
						} else {
							bobina.setPesoBrutoPrevisto(0.0d);
						}
						bobinas.add(bobina);
					}

					break;
				}

			}
			is.close();
			int totalDestinatarios = bobinas.stream()
					.collect(Collectors.groupingBy(bobina -> bobina.getNombreDestinatario(), Collectors.counting()))
					.size();
			var pesoTotal = bobinas.stream().mapToDouble(Bobina::getPesoBrutoPrevisto).sum();
			bobinasTemplate.setTotalDestinatarios(totalDestinatarios);
			bobinasTemplate.setTotalBobinas(bobinas.size());
			bobinasTemplate.setTotalPeso(pesoTotal);
			bobinasTemplate.setBobinasList(bobinas);
		} catch (Exception ex) {
			log.error("Error leyendo excel", ex);
			bobinas.clear();
		}

		return bobinasTemplate;
	}

	private String obtenerBarco(BobinasTemplate bobinasTemplate) {
		var encabezadoArray = StringUtils.split(bobinasTemplate.getEncabezado());
		var barco = new StringBuilder();
		for (String s : encabezadoArray) {
			if (NumberUtils.isDigits(s)) {
				break;
			}
			if (barco.length() != 0) {
				barco.append(StringUtils.SPACE);
			}
			if (!"MV".equals(s)) {
				barco.append(s);
			}
		}
		return barco.toString();
	}

	public Path obtenerPlantillaSalida(BobinasTemplate bobinasTemplate) {
		try {
			if (utilidades.crearDirectorio(salidaDirectory)) {
				log.info("Directorio {} creado", salidaDirectory);
			}
			var bobinas = bobinasTemplate.getBobinasList();
			var clienteEnum = ClientesEnum.getClienteEnum(bobinasTemplate.getCliente());
			switch (clienteEnum) {
			case THYSSEN:
				return obtenerPlantillaThyssen(bobinasTemplate, bobinas);
			default:
				return obtenerPlantillaArcelor(bobinasTemplate, bobinas);
			}
		} catch (Exception ex) {
			log.error("Error generando plantilla", ex);
			return Path.of(rootDirectory);
		}
	}

	private Path obtenerPlantillaArcelor(BobinasTemplate bobinasTemplate, List<Bobina> bobinas)
			throws IOException, ParseException, Exception, FileNotFoundException {
		List<Bobina> bobinasList = bobinasTemplate.getBobinasList().stream()
				.sorted(Comparator.comparing(Bobina::getNombreDestinatario).thenComparing(Bobina::getNumSerie))
				.collect(Collectors.toList());
		int bloques = bobinasList.size() % ARCELOR_BLOQUE_MAXIMOS != 0
				? (bobinasList.size() / ARCELOR_BLOQUE_MAXIMOS) + 1
				: bobinasList.size() / ARCELOR_BLOQUE_MAXIMOS;
		Map<Integer, List<Bobina>> bobinasMap = obtenerBobinasMap(bobinasList, bloques, ARCELOR_BLOQUE_MAXIMOS);
		InputStream inputStream = plantillaArcelor.getInputStream();
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheet = workbook.getSheetAt(1);
		var cabecerasRowPos = -1;
		var cabecerasColPos = -1;
		var destinatarioColPos = -1;
		var positionColPos = -1;
		var numSerieColPos = -1;
		var pesoBrutoColPos = -1;
		var pointerRowPos = -1;
		Iterator<Row> rowIterator = sheet.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				int columnIndex = cell.getColumnIndex();
				int rowIndex = cell.getRowIndex();
				switch (cell.getCellType()) {
				case STRING:
					String stringCellValue = cell.getStringCellValue();
					if (VAR_CAB.equals(stringCellValue)) {
						cabecerasRowPos = rowIndex;
						cabecerasColPos = columnIndex;
					}
					if (VAR_POSITION.equals(stringCellValue)) {
						pointerRowPos = rowIndex;
						positionColPos = columnIndex;
					}
					if (VAR_DESTINATARIO.equals(stringCellValue)) {
						destinatarioColPos = columnIndex;
					}
					if (VAR_NUM_SERIE.equals(stringCellValue)) {
						numSerieColPos = columnIndex;
					}
					if (VAR_PESO_BRUTO.equals(stringCellValue)) {
						pesoBrutoColPos = columnIndex;
					}
					break;
				default:
					break;
				}
			}
		}

		Row rowOrigenBobina = sheet.getRow(pointerRowPos);
		Row rowOrigenCabeceras = sheet.getRow(cabecerasRowPos);
		var controlUltimoRegistro = bobinasMap.size();
		var controlRegistros = 0;
		for (Map.Entry<Integer, List<Bobina>> bobinaMap : bobinasMap.entrySet()) {
			// header
			var isPrimerBloque = bobinaMap.getKey().intValue() == 0;
			generarCabeceraBloqueArcelor(workbook, sheet, bobinasTemplate, pointerRowPos, isPrimerBloque);
			controlUltimoRegistro--;
			var cabecerasIndexString = StringUtils.EMPTY;
			for (var indexBobinas = 0; indexBobinas < bobinaMap.getValue().size(); indexBobinas++) {
				controlRegistros++;
				Bobina bobina = bobinaMap.getValue().get(indexBobinas);
				if (indexBobinas != bobinaMap.getValue().size() - 1) {
					Row rowBobina = sheet.createRow(pointerRowPos + 1);
					copiarFormatoRowByCell(rowOrigenBobina, rowBobina, false, null, 0, 0);
				}
				sheet.getRow(pointerRowPos).getCell(positionColPos).setCellValue(controlRegistros);
				sheet.getRow(pointerRowPos).getCell(destinatarioColPos).setCellValue(bobina.getNombreDestinatario());
				sheet.getRow(pointerRowPos).getCell(numSerieColPos).setCellValue(bobina.getNumSerie());
				sheet.getRow(pointerRowPos).getCell(pesoBrutoColPos).setCellValue(bobina.getPesoBrutoPrevisto());
				pointerRowPos++;
			}
			if (bobinaMap.getValue().size() != ARCELOR_BLOQUE_MAXIMOS) {
				for (var k = 0; k <= ARCELOR_BLOQUE_MAXIMOS - bobinaMap.getValue().size(); k++) {
					Row rowVacio = sheet.createRow(pointerRowPos);
					copiarFormatoRowByCell(rowOrigenBobina, rowVacio, false, null, 0, 0);
					pointerRowPos++;
				}
			}
			generarPieBloqueArcelor(workbook, sheet, pointerRowPos, isPrimerBloque);
			if (controlUltimoRegistro != 0) {
				pointerRowPos += ARCELOR_BLOQUE_HEADER_SIZE + ARCELOR_BLOQUE_FOOTER_SIZE;
				Row rowCabecera = sheet.createRow(pointerRowPos);
				copiarFormatoRowByCell(rowOrigenCabeceras, rowCabecera, false, sheet, 0, 0);
				copiarContenidoRowByCell(rowOrigenCabeceras, rowCabecera);
				sheet.getRow(cabecerasRowPos).getCell(cabecerasColPos).setCellValue(cabecerasIndexString);
				cabecerasRowPos = rowCabecera.getRowNum();
				// ultima linea tiene que ser row bobina vacia
				Row rowBobina = sheet.createRow(++pointerRowPos);
				copiarFormatoRowByCell(rowOrigenBobina, rowBobina, false, null, 0, 0);
			}
			if (controlUltimoRegistro == 0) {
				sheet.getRow(cabecerasRowPos).getCell(cabecerasColPos).setCellValue(cabecerasIndexString);
			}
		}

		inputStream.close();

		var salidaGeneradaPath = new StringBuilder();
		salidaGeneradaPath.append(salidaDirectory);
		salidaGeneradaPath.append("/");
		salidaGeneradaPath.append("REPORT");
		salidaGeneradaPath.append(StringUtils.SPACE);
		salidaGeneradaPath.append("MUELLE");
		salidaGeneradaPath.append(StringUtils.SPACE);
		salidaGeneradaPath.append(bobinasTemplate.getBarco());
		salidaGeneradaPath.append(".xlsx");
		FileOutputStream outputStream = new FileOutputStream(salidaGeneradaPath.toString());
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();
		log.info("Excel creado correctamente: {}", salidaGeneradaPath);
		Path path = Path.of(salidaGeneradaPath.toString());
		return path;

	}

	private Path obtenerPlantillaThyssen(BobinasTemplate bobinasTemplate, List<Bobina> bobinas)
			throws IOException, ParseException, Exception, FileNotFoundException {
		Map<String, List<Bobina>> bobinasMap = obtenerBobinasMap(bobinas);
		var fechaHoy = utilidades.obtenerFechaString(Calendar.getInstance().getTime(), "dd/MM/yyyy");
		InputStream inputStream = plantillaThyssen.getInputStream();
		Workbook workbook = WorkbookFactory.create(inputStream);
		workbook.setSheetName(0, "Informe");
		Sheet sheet = workbook.getSheetAt(0);
		var destinatarioRowPos = -1;
		var destinatarioColPos = -1;
		var positionColPos = -1;
		var numSerieColPos = -1;
		var pesoBrutoColPos = -1;
		var bobinaRowPos = -1;
		Iterator<Row> rowIterator = sheet.iterator();
		while (rowIterator.hasNext()) {
			Row row = rowIterator.next();

			Iterator<Cell> cellIterator = row.cellIterator();
			while (cellIterator.hasNext()) {
				Cell cell = cellIterator.next();
				int columnIndex = cell.getColumnIndex();
				int rowIndex = cell.getRowIndex();
				switch (cell.getCellType()) {
				case STRING:
					String stringCellValue = cell.getStringCellValue();
					if (VAR_FECHA.equals(stringCellValue)) {
						cell.setCellValue(DateUtils.parseDate(fechaHoy, "dd/MM/yyyy"));
					}
					if (VAR_DESTINATARIO.equals(stringCellValue)) {
						destinatarioRowPos = rowIndex;
						destinatarioColPos = columnIndex;
					}
					if (VAR_POSITION.equals(stringCellValue)) {
						bobinaRowPos = rowIndex;
						positionColPos = columnIndex;
					}
					if (VAR_NUM_SERIE.equals(stringCellValue)) {
						bobinaRowPos = rowIndex;
						numSerieColPos = columnIndex;
					}
					if (VAR_PESO_BRUTO.equals(stringCellValue)) {
						bobinaRowPos = rowIndex;
						pesoBrutoColPos = columnIndex;
					}
					if (VAR_VESSEL.equals(stringCellValue)) {
						cell.setCellValue(bobinasTemplate.getBarco());
					}
					break;
				default:
					break;
				}
			}
		}
		Row rowOrigenBobina = sheet.getRow(bobinaRowPos);
		Row rowOrigenDest = sheet.getRow(destinatarioRowPos);
		var controlUltimoRegistro = bobinasMap.size();
		for (Map.Entry<String, List<Bobina>> bobinaMap : bobinasMap.entrySet()) {
			controlUltimoRegistro--;
			var destinatarioString = bobinaMap.getKey();
			List<Bobina> bobinasVal = bobinaMap.getValue().stream().sorted(Comparator.comparing(Bobina::getNumSerie))
					.collect(Collectors.toList());
			for (var indexBobinas = 0; indexBobinas < bobinasVal.size(); indexBobinas++) {
				Bobina bobina = bobinasVal.get(indexBobinas);
				if (indexBobinas != bobinasVal.size() - 1) {
					Row rowBobina = sheet.createRow(bobinaRowPos + 1);
					copiarFormatoRowByCell(rowOrigenBobina, rowBobina, false, null, 0, 0);
				}
				sheet.getRow(bobinaRowPos).getCell(positionColPos).setCellValue(indexBobinas + 1);
				sheet.getRow(bobinaRowPos).getCell(numSerieColPos).setCellValue(bobina.getNumSerie());
				sheet.getRow(bobinaRowPos).getCell(pesoBrutoColPos).setCellValue(bobina.getPesoBrutoPrevisto());
				bobinaRowPos++;
			}
			if (controlUltimoRegistro != 0) {
				Row rowDestinatario = sheet.createRow(bobinaRowPos);
				copiarFormatoRowByCell(rowOrigenDest, rowDestinatario, true, sheet, 0, 2);
				sheet.getRow(destinatarioRowPos).getCell(destinatarioColPos).setCellValue(destinatarioString);
				destinatarioRowPos = rowDestinatario.getRowNum();
				// ultima linea tiene que ser row bobina vacia
				Row rowBobina = sheet.createRow(++bobinaRowPos);
				copiarFormatoRowByCell(rowOrigenBobina, rowBobina, false, null, 0, 0);
			}
			if (controlUltimoRegistro == 0) {
				sheet.getRow(destinatarioRowPos).getCell(destinatarioColPos).setCellValue(destinatarioString);
			}
		}
		generarPiePaginaThyssen(workbook, sheet, bobinaRowPos);
		inputStream.close();

		var salidaGeneradaPath = new StringBuilder();
		salidaGeneradaPath.append(salidaDirectory);
		salidaGeneradaPath.append("/");
		salidaGeneradaPath.append("REPORT");
		salidaGeneradaPath.append(StringUtils.SPACE);
		salidaGeneradaPath.append("MUELLE");
		salidaGeneradaPath.append(StringUtils.SPACE);
		salidaGeneradaPath.append(bobinasTemplate.getBarco());
		salidaGeneradaPath.append(".xlsx");
		FileOutputStream outputStream = new FileOutputStream(salidaGeneradaPath.toString());
		workbook.write(outputStream);
		workbook.close();
		outputStream.close();
		log.info("Excel creado correctamente: {}", salidaGeneradaPath);
		Path path = Path.of(salidaGeneradaPath.toString());
		return path;
	}

	private void generarPiePaginaThyssen(Workbook wb, Sheet sheet, int bobinaRowPos) throws Exception {
		// format danios
		CellStyle styleDmg = wb.createCellStyle();
		Font fontDmg = wb.createFont();
		fontDmg.setFontName("Arial");
		fontDmg.setUnderline(Font.U_NONE);
		fontDmg.setFontHeightInPoints((short) 6);
		styleDmg.setFont(fontDmg);
		// format danios
		CellStyle styleInfo = wb.createCellStyle();
		Font fontInfo = wb.createFont();
		fontInfo.setFontName("Arial");
		fontInfo.setUnderline(Font.U_NONE);
		fontInfo.setFontHeightInPoints((short) 12);
		styleInfo.setFont(fontInfo);

		// Convert Image InputStream Into a Byte Array
		byte[] inputImageBytes = IOUtils.toByteArray(imgFirma.getInputStream());
		// Add Picture in the Workbook
		int inputImagePicture = wb.addPicture(inputImageBytes, Workbook.PICTURE_TYPE_JPEG);
		// Create a Drawing Container
		XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
		XSSFClientAnchor firmaAnchor = new XSSFClientAnchor();

		Row row1 = sheet.createRow(++bobinaRowPos);
		crearCellConStyle(row1, 0, styleDmg).setCellValue(1);
		crearCellConStyle(row1, 1, styleDmg).setCellValue("CIRCUITO EXTERIOR / OUTER WINDING");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 1, 2));
		crearCellConStyle(row1, 6, styleDmg).setCellValue(9);
		crearCellConStyle(row1, 7, styleDmg).setCellValue("ABOLLADURA / DENT");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 7, 12));
		Row row2 = sheet.createRow(++bobinaRowPos);
		crearCellConStyle(row2, 0, styleDmg).setCellValue(2);
		crearCellConStyle(row2, 1, styleDmg).setCellValue("CIRCUITO INTERIOR / INNER WINDING");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 1, 2));
		crearCellConStyle(row2, 6, styleDmg).setCellValue(10);
		crearCellConStyle(row2, 7, styleDmg).setCellValue("CORTES AGUJERO / CUT HOLE");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 7, 12));
		Row row3 = sheet.createRow(++bobinaRowPos);
		crearCellConStyle(row3, 0, styleDmg).setCellValue(3);
		crearCellConStyle(row3, 1, styleDmg).setCellValue("ZONA LATERAL / LATERAL ZONE");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 1, 2));
		crearCellConStyle(row3, 6, styleDmg).setCellValue(11);
		crearCellConStyle(row3, 7, styleDmg).setCellValue("FLEJES ROTOS / STRIPS BROKEN");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 7, 12));
		Row row4 = sheet.createRow(++bobinaRowPos);
		crearCellConStyle(row4, 0, styleDmg).setCellValue(4);
		crearCellConStyle(row4, 1, styleDmg).setCellValue("CANTONERA EXTERIOR / OUTER CORNER");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 1, 2));
		crearCellConStyle(row4, 6, styleDmg).setCellValue(12);
		crearCellConStyle(row4, 7, styleDmg).setCellValue("ABIERTO / OPENED");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 7, 12));
		Row row5 = sheet.createRow(++bobinaRowPos);
		crearCellConStyle(row5, 0, styleDmg).setCellValue(5);
		crearCellConStyle(row5, 1, styleDmg).setCellValue("CANTONERA INTERIOR / INNER CORNER");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 1, 2));
		crearCellConStyle(row5, 6, styleDmg).setCellValue(13);
		crearCellConStyle(row5, 7, styleDmg).setCellValue("DESPLAZADO / SHIFTED");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 7, 12));
		Row row6 = sheet.createRow(++bobinaRowPos);
		crearCellConStyle(row6, 0, styleDmg).setCellValue("");
		crearCellConStyle(row6, 1, styleDmg).setCellValue("");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 1, 2));
		crearCellConStyle(row6, 6, styleDmg).setCellValue(14);
		crearCellConStyle(row6, 7, styleDmg).setCellValue("HUMEDO / WET");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 7, 12));
		Row row7 = sheet.createRow(++bobinaRowPos);
		crearCellConStyle(row7, 0, styleDmg).setCellValue(6);
		crearCellConStyle(row7, 1, styleDmg).setCellValue("CONT DAÑADO / DAMAGED CONTENT");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 1, 2));
		crearCellConStyle(row7, 6, styleDmg).setCellValue(15);
		crearCellConStyle(row7, 7, styleDmg).setCellValue("OXIDO / RUSTY");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 7, 12));
		Row row8 = sheet.createRow(++bobinaRowPos);
		crearCellConStyle(row8, 0, styleDmg).setCellValue(7);
		crearCellConStyle(row8, 1, styleDmg).setCellValue("DEFORMA OVAL / OVAL DEFORMATION");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 1, 2));
		crearCellConStyle(row8, 6, styleDmg).setCellValue(16);
		crearCellConStyle(row8, 7, styleDmg).setCellValue("FOTO / PHOTO");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 7, 12));
		crearCellConStyle(row8, 14, styleInfo).setCellValue("Jose Menduiña");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 14, 18));
		Row row9 = sheet.createRow(++bobinaRowPos);
		crearCellConStyle(row9, 0, styleDmg).setCellValue(8);
		crearCellConStyle(row9, 1, styleDmg).setCellValue("TELESCOPICA / TELESCOPIC");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 1, 2));
		crearCellConStyle(row9, 6, styleDmg).setCellValue(17);
		crearCellConStyle(row9, 7, styleDmg).setCellValue("DISCK Nº");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 7, 12));
		crearCellConStyle(row9, 14, styleInfo).setCellValue("ATM OCA GLOBAL");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 14, 18));
		Row row10 = sheet.createRow(++bobinaRowPos);
		crearCellConStyle(row10, 0, styleInfo).setCellValue("");
		crearCellConStyle(row10, 1, styleInfo).setCellValue("All the damages are checked into the hold.");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 1, 12));
		Row row11 = sheet.createRow(++bobinaRowPos);
		crearCellConStyle(row11, 0, styleInfo).setCellValue("");
		crearCellConStyle(row11, 1, styleInfo).setCellValue("Agent");
		crearCellConStyle(row11, 6, styleInfo).setCellValue("Stevedor");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 6, 9));
		crearCellConStyle(row11, 13, styleInfo).setCellValue("Captain");
		sheet.addMergedRegion(new CellRangeAddress(bobinaRowPos, bobinaRowPos, 13, 15));
		// 4x4
		firmaAnchor.setCol1(14); // Sets the column (0 based) of the first cell.
		firmaAnchor.setCol2(18); // Sets the column (0 based) of the Second cell.
		firmaAnchor.setRow1(row4.getRowNum()); // Sets the row (0 based) of the first cell.
		firmaAnchor.setRow2(row8.getRowNum()); // Sets the row (0 based) of the Second cell.
		//
		drawing.createPicture(firmaAnchor, inputImagePicture);

	}

	private Cell crearCellConStyle(Row row, int cellNum, CellStyle cellStyle) {
		Cell cell = row.createCell(cellNum);
		cell.setCellStyle(cellStyle);
		return cell;
	}

	private Map<String, List<Bobina>> obtenerBobinasMap(List<Bobina> bobinas) {
		Map<String, List<Bobina>> bobinasMap = new HashMap<String, List<Bobina>>();
		var destinatarios = new ArrayList<String>();
		bobinas.forEach(bobina -> {
			if (!destinatarios.contains(bobina.getNombreDestinatario())) {
				destinatarios.add(bobina.getNombreDestinatario());
			}
		});
		for (var destinatario : destinatarios) {
			List<Bobina> bobinaDatos = bobinas.stream()
					.filter(bobina -> bobina.getNombreDestinatario().equals(destinatario)).collect(Collectors.toList());
			bobinasMap.put(destinatario, bobinaDatos);
		}
		Map<String, List<Bobina>> bobinasMapSorted = new TreeMap<String, List<Bobina>>(bobinasMap);
		return bobinasMapSorted;
	}

	private void copiarFormatoRowByCell(Row rowOrigen, Row rowDestino, boolean merged, Sheet sheet, int mergedStars,
			int mergedEnds) {
		var rowDestinoPos = rowDestino.getRowNum();
		Iterator<Cell> cellIterator = rowOrigen.iterator();
		var i = 0;
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			rowDestino.createCell(i).setCellStyle(cell.getCellStyle());
			i++;
		}
		if (merged) {
			sheet.addMergedRegion(new CellRangeAddress(rowDestinoPos, rowDestinoPos, mergedStars, mergedEnds));
		}

	}

	private void copiarContenidoRowByCell(Row rowOrigen, Row rowDestino) {
		Iterator<Cell> cellIterator = rowOrigen.iterator();
		var i = 0;
		while (cellIterator.hasNext()) {
			Cell cell = cellIterator.next();
			switch (cell.getCellType()) {
			case STRING:
				rowDestino.getCell(i).setCellValue(cell.getStringCellValue());
				break;
			case NUMERIC:
				rowDestino.getCell(i).setCellValue(cell.getNumericCellValue());
				break;
			default:
				rowDestino.getCell(i).setCellValue(StringUtils.EMPTY);
				break;
			}
			i++;
		}

	}

	private Map<Integer, List<Bobina>> obtenerBobinasMap(List<Bobina> bobinas, int bloques, int bloqueMax) {
		var bobinasMap = new HashMap<Integer, List<Bobina>>();
		var bobinasIndex = 0;
		for (var i = 0; i < bloques; i++) {
			List<Bobina> bobinasBloque = new ArrayList<Bobina>();
			for (var j = 0; bobinasIndex < bobinas.size() && j < bloqueMax; j++) {
				bobinasBloque.add(bobinas.get(bobinasIndex));
				bobinasIndex++;
			}
			bobinasMap.put(i, bobinasBloque);
		}
		return new TreeMap<Integer, List<Bobina>>(bobinasMap);
	}

	private void generarPieBloqueArcelor(final Workbook workbook, final Sheet sheet, final int pointerRowPos,
			final boolean isPrimerBloque) {
		//
		var pointerStartHeader = pointerRowPos;
		var panel1col1 = 0;
		var panel1row1 = 57;
		var panel2col1 = 0;
		var panel2row1 = 65;
		//
		HSSFPatriarch pat = (HSSFPatriarch) sheet.getDrawingPatriarch();

		if (!isPrimerBloque) {
			int lineStyleColor = 0;
			HSSFRichTextString panel1RichString = new HSSFRichTextString(StringUtils.EMPTY);
			HSSFRichTextString panel2RichString = new HSSFRichTextString(StringUtils.EMPTY);
			for (HSSFShape shape : pat.getChildren()) {
				var clienteAnchor = (HSSFClientAnchor) shape.getAnchor();
				if (shape instanceof HSSFTextbox) {
					HSSFTextbox textbox = (HSSFTextbox) shape;
					HSSFRichTextString richString = textbox.getString();
					lineStyleColor = textbox.getLineStyleColor();
					String contenidoString = richString.getString();
					if (panel1col1 == clienteAnchor.getCol1() && panel1row1 == clienteAnchor.getRow1()) {
						panel1RichString = new HSSFRichTextString(contenidoString);
						panel1RichString.applyFont(richString.getFontAtIndex(0));
					}
					if (panel2col1 == clienteAnchor.getCol1() && panel2row1 == clienteAnchor.getRow1()) {
						panel2RichString = new HSSFRichTextString(contenidoString);

						var stringSplit = StringUtils.split(contenidoString);
						for (String s : stringSplit) {
							int startIndex = contenidoString.indexOf(s);
							int endIndex = startIndex + s.length();
							panel2RichString.applyFont(startIndex, endIndex, richString.getFontAtIndex(0));
						}
					}

				}
			}
			HSSFTextbox shapePanel1 = pat.createTextbox(new HSSFClientAnchor(0, 25, 0, 0, (short) 1, pointerStartHeader,
					(short) 23, pointerStartHeader + 8));
			shapePanel1.setLineStyleColor(lineStyleColor);
			shapePanel1.setString(panel1RichString);
			shapePanel1.setMarginTop(MARGIN_PANELES_TEXT);
			shapePanel1.setMarginBottom(MARGIN_PANELES_TEXT);
			shapePanel1.setMarginLeft(MARGIN_PANELES_TEXT);
			shapePanel1.setMarginRight(MARGIN_PANELES_TEXT);
			pointerStartHeader += 8;
			HSSFTextbox shapePanel2 = pat.createTextbox(new HSSFClientAnchor(0, 20, 0, 0, (short) 1, pointerStartHeader,
					(short) 23, pointerStartHeader + 6));
			shapePanel2.setLineStyleColor(lineStyleColor);
			shapePanel2.setString(panel2RichString);
			shapePanel2.setMarginTop(MARGIN_PANELES_TEXT);
			shapePanel2.setMarginBottom(MARGIN_PANELES_TEXT);
			shapePanel2.setMarginLeft(MARGIN_PANELES_TEXT);
			shapePanel2.setMarginRight(MARGIN_PANELES_TEXT);
		}

	}

	public void generarCabeceraBloqueArcelor(final Workbook workbook, final Sheet sheet,
			BobinasTemplate bobinasTemplate, final int pointerRowPos, boolean isPrimerBloque) {
		var pointerStartHeader = pointerRowPos - ARCELOR_BLOQUE_HEADER_SIZE - 1;
		var fechaHoy = utilidades.obtenerFechaString(Calendar.getInstance().getTime(), "dd/MM/yy");
		Font fontDatosInsert = workbook.createFont();
		fontDatosInsert.setFontName("Calibri");
		fontDatosInsert.setFontHeightInPoints((short) 9);
		fontDatosInsert.setBold(true);
		//
		var panel1col1 = 12;
		var panel1row1 = 0;
		var panel2col1 = 0;
		var panel2row1 = 4;
		var panel3col1 = 0;
		var panel3row1 = 8;
		//
		HSSFPatriarch pat = (HSSFPatriarch) sheet.getDrawingPatriarch();
		if (!isPrimerBloque) {
			int pictureIndex = -1;
			for (HSSFShape shape : pat.getChildren()) {
				if (shape instanceof HSSFPicture) {
					HSSFPicture picture = (HSSFPicture) shape;
					pictureIndex = picture.getPictureIndex();
					break;
				}
			}
			HSSFPicture shapePicture = pat.createPicture(
					new HSSFClientAnchor(0, 0, 0, 0, (short) 1, pointerStartHeader, (short) 2, pointerStartHeader + 4),
					pictureIndex);
			shapePicture.resize(1.0, 1.0);
			int lineStyleColor = 0;
			HSSFRichTextString panel1RichString = new HSSFRichTextString(StringUtils.EMPTY);
			HSSFRichTextString panel2RichString = new HSSFRichTextString(StringUtils.EMPTY);
			HSSFRichTextString panel3RichString = new HSSFRichTextString(StringUtils.EMPTY);
			for (HSSFShape shape : pat.getChildren()) {
				var clienteAnchor = (HSSFClientAnchor) shape.getAnchor();
				if (shape instanceof HSSFTextbox) {
					HSSFTextbox textbox = (HSSFTextbox) shape;
					HSSFRichTextString richString = textbox.getString();
					lineStyleColor = textbox.getLineStyleColor();
					if (panel1col1 == clienteAnchor.getCol1() && panel1row1 == clienteAnchor.getRow1()) {
						var s = richString.getString();
						panel1RichString = new HSSFRichTextString(richString.getString());
						panel1RichString.applyFont(richString.getFontOfFormattingRun(0));
						panel1RichString.applyFont(s.indexOf("HI"), s.length(), richString.getFontOfFormattingRun(1));
					}
					// datos barco fecha
					if (panel2col1 == clienteAnchor.getCol1() && panel2row1 == clienteAnchor.getRow1()) {
						var s = richString.getString();
						panel2RichString = new HSSFRichTextString(richString.getString());
						panel2RichString.applyFont(richString.getFontAtIndex(0));
						int startIndex = s.indexOf(fechaHoy);
						int endIndex = startIndex + fechaHoy.length();
						panel2RichString.applyFont(startIndex, endIndex, richString.getFontOfFormattingRun(1));
						startIndex = s.indexOf(bobinasTemplate.getBarco());
						endIndex = startIndex + bobinasTemplate.getBarco().length();
						panel2RichString.applyFont(startIndex, endIndex, richString.getFontOfFormattingRun(1));
					}
					if (panel3col1 == clienteAnchor.getCol1() && panel3row1 == clienteAnchor.getRow1()) {
						panel3RichString = new HSSFRichTextString(richString.getString());
						panel3RichString.applyFont(richString.getFontAtIndex(0));
					}
				}
			}
			HSSFTextbox shapePanel1 = pat.createTextbox(new HSSFClientAnchor(0, 0, 0, 0, (short) panel1col1,
					pointerStartHeader, (short) (panel1col1 + 11), pointerStartHeader + 4));
			shapePanel1.setLineStyleColor(lineStyleColor);
			shapePanel1.setString(panel1RichString);
			shapePanel1.setMarginTop(MARGIN_PANELES_TEXT);
			shapePanel1.setMarginBottom(MARGIN_PANELES_TEXT);
			shapePanel1.setMarginLeft(MARGIN_PANELES_TEXT);
			shapePanel1.setMarginRight(MARGIN_PANELES_TEXT);
			pointerStartHeader += 4;
			HSSFTextbox shapePanel2 = pat.createTextbox(new HSSFClientAnchor(0, 20, 0, 0, (short) 1, pointerStartHeader,
					(short) 5, pointerStartHeader + 4));
			shapePanel2.setLineStyleColor(lineStyleColor);
			shapePanel2.setString(panel2RichString);
			shapePanel2.setMarginTop(MARGIN_PANELES_TEXT);
			shapePanel2.setMarginBottom(MARGIN_PANELES_TEXT);
			shapePanel2.setMarginLeft(MARGIN_PANELES_TEXT);
			shapePanel2.setMarginRight(MARGIN_PANELES_TEXT);

			pointerStartHeader += 4;
			HSSFTextbox shapePanel3 = pat.createTextbox(new HSSFClientAnchor(0, 15, 0, 200, (short) 1,
					pointerStartHeader, (short) 5, pointerStartHeader + 4));
			shapePanel3.setLineStyleColor(lineStyleColor);
			shapePanel3.setString(panel3RichString);
			shapePanel3.setMarginTop(MARGIN_PANELES_TEXT);
			shapePanel3.setMarginBottom(MARGIN_PANELES_TEXT);
			shapePanel3.setMarginLeft(MARGIN_PANELES_TEXT);
			shapePanel3.setMarginRight(MARGIN_PANELES_TEXT);

		}
		if (isPrimerBloque) {
			for (HSSFShape shape : pat.getChildren()) {
				if (shape instanceof HSSFTextbox) {
					HSSFTextbox textbox = (HSSFTextbox) shape;
					textbox.setMarginBottom(MARGIN_PANELES_TEXT);
					textbox.setMarginTop(MARGIN_PANELES_TEXT);
					textbox.setMarginRight(MARGIN_PANELES_TEXT);
					textbox.setMarginLeft(MARGIN_PANELES_TEXT);
					HSSFRichTextString richString = textbox.getString();
					String str = richString.getString();
					if (str.contains(VAR_VESSEL) && str.contains(VAR_FECHA)) {
						var strUpdate = StringUtils.replace(str, VAR_VESSEL, bobinasTemplate.getBarco());
						strUpdate = StringUtils.replace(strUpdate, VAR_FECHA, fechaHoy);
						// update
						HSSFRichTextString stringRichUpdate = new HSSFRichTextString(strUpdate);
						stringRichUpdate.applyFont(richString.getFontAtIndex(0));

						int startIndex = strUpdate.indexOf(fechaHoy);
						int endIndex = startIndex + fechaHoy.length();
						stringRichUpdate.applyFont(startIndex, endIndex, fontDatosInsert);
						startIndex = strUpdate.indexOf(bobinasTemplate.getBarco());
						endIndex = startIndex + bobinasTemplate.getBarco().length();
						stringRichUpdate.applyFont(startIndex, endIndex, fontDatosInsert);
						textbox.setString(stringRichUpdate);

					}
				}
			}
		}
	}

}
