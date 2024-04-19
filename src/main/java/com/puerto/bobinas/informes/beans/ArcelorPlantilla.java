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
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextParagraph;
import org.apache.poi.xssf.usermodel.XSSFTextRun;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Component;

import com.puerto.bobinas.informes.enums.ClientesEnum;
import com.puerto.bobinas.informes.utils.Utilidades;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class ArcelorPlantilla extends AbstractPlantilla {

	@Autowired
	private Utilidades utilidades;

	@Value("${fileChooser.directory.excels.salida}")
	private String salidaDirectory;
	@Value("classpath:plantillas/plantilla_arcelor.xlsx")
	private Resource plantillaArcelor;
	private static final int ARCELOR_BLOQUE_HEADER_SIZE = 13;
	private static final int ARCELOR_BLOQUE_FOOTER_SIZE = 15;
	private static final int ARCELOR_BLOQUE_MAXIMOS = 43;

	@Override
	public Path generarPlantilla(BobinasTemplate bobinasTemplate) throws Exception {
		List<Bobina> bobinasList = bobinasTemplate.getBobinasList().stream()
				.sorted(Comparator.comparing(Bobina::getNombreDestinatario).thenComparing(Bobina::getNumSerie))
				.collect(Collectors.toList());
		int bloques = bobinasList.size() % ARCELOR_BLOQUE_MAXIMOS != 0
				? (bobinasList.size() / ARCELOR_BLOQUE_MAXIMOS) + 1
				: bobinasList.size() / ARCELOR_BLOQUE_MAXIMOS;
		Map<Object, List<Bobina>> bobinasMap = obtenerBobinasMap(bobinasList, bloques, ARCELOR_BLOQUE_MAXIMOS);
		InputStream inputStream = plantillaArcelor.getInputStream();
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheet = workbook.getSheetAt(0);
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
		for (Map.Entry<Object, List<Bobina>> bobinaMap : bobinasMap.entrySet()) {
			// header
			Integer key = (Integer) bobinaMap.getKey();
			var isPrimerBloque = key.intValue() == 0;
			generarCabeceraBloque(workbook, sheet, bobinasTemplate, pointerRowPos, isPrimerBloque);
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
			generarPieBloque(workbook, sheet, pointerRowPos, isPrimerBloque);
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
		salidaGeneradaPath.append(ClientesEnum.ARCELOR.name());
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
		var pointerStartHeader = pointerRowPos - ARCELOR_BLOQUE_HEADER_SIZE - 1;
		var fechaHoy = utilidades.obtenerFechaString(Calendar.getInstance().getTime(), "dd/MM/yy");
		XSSFFont font = (XSSFFont) workbook.createFont();
		//
		var panel1col1 = 1;
		var panel1row1 = 3;
		var panel2col1 = 1;
		var panel2row1 = 7;
		//
		XSSFDrawing pat = (XSSFDrawing) sheet.getDrawingPatriarch();
		if (isPrimerBloque) {
			for (XSSFShape shape : pat.getShapes()) {
				if (shape instanceof XSSFSimpleShape) {
					XSSFSimpleShape simpleShape = (XSSFSimpleShape) shape;
					String text = simpleShape.getText();
					if (StringUtils.contains(text, VAR_VESSEL)) {
						XSSFRichTextString richText = new XSSFRichTextString();
						List<XSSFTextParagraph> textParagraphs = simpleShape.getTextParagraphs();
						for (XSSFTextParagraph textParagraph : textParagraphs) {
							for (XSSFTextRun xssfTextRun : textParagraph.getTextRuns()) {
								String textRun = xssfTextRun.getText();
								font.setFontName(xssfTextRun.getFontFamily());
								font.setFontHeight(xssfTextRun.getFontSize());
								font.setBold(xssfTextRun.isBold());
								if (textRun.contains(VAR_VESSEL)) {
									richText.append(
											StringUtils.replace(textRun, VAR_VESSEL, bobinasTemplate.getBarco()), font);
								} else if (textRun.contains(VAR_FECHA)) {
									richText.append(StringUtils.replace(textRun, VAR_FECHA, fechaHoy), font);
								} else {
									richText.append(textRun, font);
								}
							}
							richText.append(StringUtils.LF);
						}
						simpleShape.setText(richText);
					}
				}
			}
		}
		if (!isPrimerBloque) {
			for (XSSFShape shape : pat.getShapes()) {
				if (shape instanceof XSSFPicture) {
					copyPicture(shape, (XSSFSheet) sheet, pat, -1, -1, pointerStartHeader, pointerStartHeader + 3);
					break;
				}
			}
			XSSFShape panel1Shape = null;
			XSSFShape panel2Shape = null;
			for (XSSFShape shape : pat.getShapes()) {
				if (shape instanceof XSSFSimpleShape) {
					XSSFClientAnchor anchor = (XSSFClientAnchor) shape.getAnchor();
					if (anchor.getCol1() == panel1col1 && anchor.getRow1() == panel1row1) {
						panel1Shape = shape;
					}
					if (anchor.getCol1() == panel2col1 && anchor.getRow1() == panel2row1) {
						panel2Shape = shape;
					}
				}
			}
			pointerStartHeader += 3;
			copySimpleShape(panel1Shape, (XSSFSheet) sheet, pat, -1, -1, pointerStartHeader, pointerStartHeader + 4,
					true);
			pointerStartHeader += 4;
			copySimpleShape(panel2Shape, (XSSFSheet) sheet, pat, -1, -1, pointerStartHeader, pointerStartHeader + 5,
					true);
		}

	}

	@Override
	protected void generarPieBloque(Workbook workbook, Sheet sheet, int pointerRowPos, boolean isPrimerBloque)
			throws IOException {
		//
		var pointerStartHeader = pointerRowPos;
		var panel1col1 = 0;
		var panel1row1 = 57;
		var panel2col1 = 0;
		var panel2row1 = 65;
		//
		XSSFDrawing pat = (XSSFDrawing) sheet.getDrawingPatriarch();
		if (!isPrimerBloque) {
			XSSFShape panel1Shape = null;
			XSSFShape panel2Shape = null;
			for (XSSFShape shape : pat.getShapes()) {
				if (shape instanceof XSSFSimpleShape) {
					XSSFClientAnchor anchor = (XSSFClientAnchor) shape.getAnchor();
					if (anchor.getCol1() == panel1col1 && anchor.getRow1() == panel1row1) {
						panel1Shape = shape;
					}
					if (anchor.getCol1() == panel2col1 && anchor.getRow1() == panel2row1) {
						panel2Shape = shape;
					}
				}
			}
			copySimpleShape(panel1Shape, (XSSFSheet) sheet, pat, -1, 21, pointerStartHeader, pointerStartHeader + 8,
					true);
			pointerStartHeader += 8;
			copySimpleShape(panel2Shape, (XSSFSheet) sheet, pat, -1, 21, pointerStartHeader, pointerStartHeader + 6,
					true);
		}

	}

	@Override
	protected Map<Object, List<Bobina>> obtenerBobinasMap(List<Bobina> bobinas, int bloques, int bloqueMax) {
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
		return new TreeMap<Object, List<Bobina>>(bobinasMap);
	}

}
