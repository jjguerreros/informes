package com.puerto.bobinas.informes.beans;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Comparator;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
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

import com.puerto.bobinas.informes.enums.ClientesEnum;
import com.puerto.bobinas.informes.utils.Utilidades;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class ThyssenPlantilla extends AbstractPlantilla {

	@Autowired
	private Utilidades utilidades;

	@Value("classpath:img/firma.jpg")
	private Resource imgFirma;
	@Value("classpath:plantillas/plantilla_thyssen.xlsx")
	private Resource plantillaResource;
	@Value("${fileChooser.directory.excels.salida}")
	private String salidaDirectory;

	@Override
	public Path generarPlantilla(BobinasTemplate bobinasTemplate) throws Exception {
		var bobinas = bobinasTemplate.getBobinasList();
		Map<Object, List<Bobina>> bobinasMap = obtenerBobinasMap(bobinas, 0, 0);
		var fechaHoy = utilidades.obtenerFechaString(Calendar.getInstance().getTime(), "dd/MM/yyyy");
		InputStream inputStream = plantillaResource.getInputStream();
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
		for (Map.Entry<Object, List<Bobina>> bobinaMap : bobinasMap.entrySet()) {
			controlUltimoRegistro--;
			var destinatarioString = (String) bobinaMap.getKey();
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
		generarPieBloque(workbook, sheet, bobinaRowPos, false);
		inputStream.close();

		var salidaGeneradaPath = new StringBuilder();
		salidaGeneradaPath.append(salidaDirectory);
		salidaGeneradaPath.append("/");
		salidaGeneradaPath.append(ClientesEnum.THYSSEN.name());
		if (utilidades.crearDirectorio(salidaGeneradaPath.toString())) {
			log.info("Directorio {} creado", Path.of(salidaGeneradaPath.toString()));
		}
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
		log.info("Excel creado correctamente: {}", Path.of(salidaGeneradaPath.toString()));
		Path path = Path.of(salidaGeneradaPath.toString());
		return path;
	}

	@Override
	protected void generarCabeceraBloque(Workbook workbook, Sheet sheet, BobinasTemplate bobinasTemplate,
			int pointerRowPos, boolean isPrimerBloque) {
	}

	@Override
	protected void generarPieBloque(Workbook workbook, Sheet sheet, int pointerRowPos, boolean isPrimerBloque)
			throws IOException {
		// format danios
		CellStyle styleDmg = workbook.createCellStyle();
		Font fontDmg = workbook.createFont();
		fontDmg.setFontName("Arial");
		fontDmg.setUnderline(Font.U_NONE);
		fontDmg.setFontHeightInPoints((short) 6);
		styleDmg.setFont(fontDmg);
		// format danios
		CellStyle styleInfo = workbook.createCellStyle();
		Font fontInfo = workbook.createFont();
		fontInfo.setFontName("Arial");
		fontInfo.setUnderline(Font.U_NONE);
		fontInfo.setFontHeightInPoints((short) 12);
		styleInfo.setFont(fontInfo);

		// Convert Image InputStream Into a Byte Array
		byte[] inputImageBytes = IOUtils.toByteArray(imgFirma.getInputStream());
		// Add Picture in the Workbook
		int inputImagePicture = workbook.addPicture(inputImageBytes, Workbook.PICTURE_TYPE_JPEG);
		// Create a Drawing Container
		XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();
		XSSFClientAnchor firmaAnchor = new XSSFClientAnchor();

		Row row1 = sheet.createRow(++pointerRowPos);
		crearCellConStyle(row1, 0, styleDmg).setCellValue(1);
		crearCellConStyle(row1, 1, styleDmg).setCellValue("CIRCUITO EXTERIOR / OUTER WINDING");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 1, 2));
		crearCellConStyle(row1, 6, styleDmg).setCellValue(9);
		crearCellConStyle(row1, 7, styleDmg).setCellValue("ABOLLADURA / DENT");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 7, 12));
		Row row2 = sheet.createRow(++pointerRowPos);
		crearCellConStyle(row2, 0, styleDmg).setCellValue(2);
		crearCellConStyle(row2, 1, styleDmg).setCellValue("CIRCUITO INTERIOR / INNER WINDING");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 1, 2));
		crearCellConStyle(row2, 6, styleDmg).setCellValue(10);
		crearCellConStyle(row2, 7, styleDmg).setCellValue("CORTES AGUJERO / CUT HOLE");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 7, 12));
		Row row3 = sheet.createRow(++pointerRowPos);
		crearCellConStyle(row3, 0, styleDmg).setCellValue(3);
		crearCellConStyle(row3, 1, styleDmg).setCellValue("ZONA LATERAL / LATERAL ZONE");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 1, 2));
		crearCellConStyle(row3, 6, styleDmg).setCellValue(11);
		crearCellConStyle(row3, 7, styleDmg).setCellValue("FLEJES ROTOS / STRIPS BROKEN");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 7, 12));
		Row row4 = sheet.createRow(++pointerRowPos);
		crearCellConStyle(row4, 0, styleDmg).setCellValue(4);
		crearCellConStyle(row4, 1, styleDmg).setCellValue("CANTONERA EXTERIOR / OUTER CORNER");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 1, 2));
		crearCellConStyle(row4, 6, styleDmg).setCellValue(12);
		crearCellConStyle(row4, 7, styleDmg).setCellValue("ABIERTO / OPENED");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 7, 12));
		Row row5 = sheet.createRow(++pointerRowPos);
		crearCellConStyle(row5, 0, styleDmg).setCellValue(5);
		crearCellConStyle(row5, 1, styleDmg).setCellValue("CANTONERA INTERIOR / INNER CORNER");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 1, 2));
		crearCellConStyle(row5, 6, styleDmg).setCellValue(13);
		crearCellConStyle(row5, 7, styleDmg).setCellValue("DESPLAZADO / SHIFTED");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 7, 12));
		Row row6 = sheet.createRow(++pointerRowPos);
		crearCellConStyle(row6, 0, styleDmg).setCellValue("");
		crearCellConStyle(row6, 1, styleDmg).setCellValue("");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 1, 2));
		crearCellConStyle(row6, 6, styleDmg).setCellValue(14);
		crearCellConStyle(row6, 7, styleDmg).setCellValue("HUMEDO / WET");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 7, 12));
		Row row7 = sheet.createRow(++pointerRowPos);
		crearCellConStyle(row7, 0, styleDmg).setCellValue(6);
		crearCellConStyle(row7, 1, styleDmg).setCellValue("CONT DAÑADO / DAMAGED CONTENT");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 1, 2));
		crearCellConStyle(row7, 6, styleDmg).setCellValue(15);
		crearCellConStyle(row7, 7, styleDmg).setCellValue("OXIDO / RUSTY");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 7, 12));
		Row row8 = sheet.createRow(++pointerRowPos);
		crearCellConStyle(row8, 0, styleDmg).setCellValue(7);
		crearCellConStyle(row8, 1, styleDmg).setCellValue("DEFORMA OVAL / OVAL DEFORMATION");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 1, 2));
		crearCellConStyle(row8, 6, styleDmg).setCellValue(16);
		crearCellConStyle(row8, 7, styleDmg).setCellValue("FOTO / PHOTO");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 7, 12));
		crearCellConStyle(row8, 14, styleInfo).setCellValue("Jose Menduiña");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 14, 18));
		Row row9 = sheet.createRow(++pointerRowPos);
		crearCellConStyle(row9, 0, styleDmg).setCellValue(8);
		crearCellConStyle(row9, 1, styleDmg).setCellValue("TELESCOPICA / TELESCOPIC");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 1, 2));
		crearCellConStyle(row9, 6, styleDmg).setCellValue(17);
		crearCellConStyle(row9, 7, styleDmg).setCellValue("DISCK Nº");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 7, 12));
		crearCellConStyle(row9, 14, styleInfo).setCellValue("ATM OCA GLOBAL");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 14, 18));
		Row row10 = sheet.createRow(++pointerRowPos);
		crearCellConStyle(row10, 0, styleInfo).setCellValue("");
		crearCellConStyle(row10, 1, styleInfo).setCellValue("All the damages are checked into the hold.");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 1, 12));
		Row row11 = sheet.createRow(++pointerRowPos);
		crearCellConStyle(row11, 0, styleInfo).setCellValue("");
		crearCellConStyle(row11, 1, styleInfo).setCellValue("Agent");
		crearCellConStyle(row11, 6, styleInfo).setCellValue("Stevedor");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 6, 9));
		crearCellConStyle(row11, 13, styleInfo).setCellValue("Captain");
		sheet.addMergedRegion(new CellRangeAddress(pointerRowPos, pointerRowPos, 13, 15));
		// 4x4
		firmaAnchor.setCol1(14); // Sets the column (0 based) of the first cell.
		firmaAnchor.setCol2(18); // Sets the column (0 based) of the Second cell.
		firmaAnchor.setRow1(row4.getRowNum()); // Sets the row (0 based) of the first cell.
		firmaAnchor.setRow2(row8.getRowNum()); // Sets the row (0 based) of the Second cell.
		//
		drawing.createPicture(firmaAnchor, inputImagePicture);

	}

	@Override
	protected Map<Object, List<Bobina>> obtenerBobinasMap(List<Bobina> bobinas, int bloques, int bloqueMax) {
		Map<Object, List<Bobina>> bobinasMap = new HashMap<Object, List<Bobina>>();
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
		Map<Object, List<Bobina>> bobinasMapSorted = new TreeMap<Object, List<Bobina>>(bobinasMap);
		return bobinasMapSorted;
	}

}
