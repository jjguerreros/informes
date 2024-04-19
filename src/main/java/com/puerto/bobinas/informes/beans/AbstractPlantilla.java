package com.puerto.bobinas.informes.beans;

import java.io.IOException;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFCreationHelper;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextParagraph;
import org.apache.poi.xssf.usermodel.XSSFTextRun;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape;

public abstract class AbstractPlantilla {

	protected final String VAR_CAB = "VAR_CAB";

	protected final String VAR_PESO_BRUTO = "VAR_PESO_BRUTO";

	protected final String VAR_NUM_SERIE = "VAR_NUM_SERIE";

	protected final String VAR_POSITION = "VAR_POSITION";

	protected final String VAR_DESTINATARIO = "VAR_DESTINATARIO";

	protected final String VAR_VESSEL = "VAR_VESSEL";

	protected final String VAR_FECHA = "VAR_FECHA";

	public abstract Path generarPlantilla(final BobinasTemplate bobinasTemplate) throws Exception;

	protected abstract void generarCabeceraBloque(final Workbook workbook, final Sheet sheet,
			BobinasTemplate bobinasTemplate, final int pointerRowPos, boolean isPrimerBloque);

	protected abstract void generarPieBloque(final Workbook workbook, final Sheet sheet, final int pointerRowPos,
			final boolean isPrimerBloque) throws IOException;

	abstract protected Map<Object, List<Bobina>> obtenerBobinasMap(List<Bobina> bobinas, int bloques, int bloqueMax);

	protected void copySimpleShape(XSSFShape shape, XSSFSheet sheet, XSSFDrawing pat, int col1, int col2, int row1,
			int row2, boolean copyChoord) {
		XSSFSimpleShape simpleShape = (XSSFSimpleShape) shape;
		XSSFClientAnchor anchor = (XSSFClientAnchor) shape.getAnchor();

		XSSFWorkbook wb = sheet.getWorkbook();
		XSSFCreationHelper newHelper = wb.getCreationHelper();
		XSSFClientAnchor newAnchor = newHelper.createClientAnchor();
		// Row / Column placement.
		if (col1 == -1) {
			newAnchor.setCol1(anchor.getCol1());
		} else {
			newAnchor.setCol1(col1);
		}
		if (col2 == -1) {
			newAnchor.setCol2(anchor.getCol2());
		} else {
			newAnchor.setCol2(col2);
		}
		newAnchor.setRow1(row1);
		newAnchor.setRow2(row2);

		// Fine touch adjustment along the XY coordinate.
		if (copyChoord) {
			newAnchor.setDx1(anchor.getDx1());
			newAnchor.setDx2(anchor.getDx2());
			newAnchor.setDy1(anchor.getDy1());
			newAnchor.setDy2(anchor.getDy2());
		}
		var isUnderline = false;
		var underLineList = new ArrayList<String>();
		XSSFFont font = (XSSFFont) sheet.getWorkbook().createFont();
		XSSFRichTextString richText = new XSSFRichTextString();
		List<XSSFTextParagraph> textParagraphs = simpleShape.getTextParagraphs();
		for (XSSFTextParagraph textParagraph : textParagraphs) {
			for (XSSFTextRun xssfTextRun : textParagraph.getTextRuns()) {
				String textRun = xssfTextRun.getText();
				font.setFontName(xssfTextRun.getFontFamily());
				font.setFontHeight(xssfTextRun.getFontSize());
				font.setBold(xssfTextRun.isBold());
				if (xssfTextRun.isUnderline()) {
					isUnderline = true;
					underLineList.add(StringUtils.trim(textRun));
					richText.append(textRun);
				} else {
					richText.append(textRun, font);
				}
			}
			richText.append(StringUtils.LF);
		}
		if (isUnderline) {
			font.setUnderline(XSSFFont.U_SINGLE);
			for (var s : underLineList) {
				int iStarts = richText.getString().indexOf(s);
				int iEnds = iStarts + s.length();
				richText.applyFont(iStarts, iEnds, font);
			}
		}
		XSSFSimpleShape simpleShapeCreated = pat.createSimpleShape(newAnchor);
		simpleShapeCreated.setText(richText);
		CTShape ctShape = simpleShape.getCTShape();
		CTShapeProperties ctShapeProperties = ctShape.getSpPr();
		// background
		simpleShapeCreated.getCTShape().getSpPr().setSolidFill(ctShapeProperties.getSolidFill());
		// borde
		simpleShapeCreated.getCTShape().getSpPr().setLn(ctShapeProperties.getLn());
		;

	}

	protected void copyPicture(XSSFShape shape, XSSFSheet sheet, XSSFDrawing pat, int col1, int col2, int row1,
			int row2) {
		XSSFPicture picture = (XSSFPicture) shape;

		XSSFPictureData xssfPictureData = picture.getPictureData();
		XSSFClientAnchor anchor = (XSSFClientAnchor) shape.getAnchor();

		int x1 = anchor.getDx1();
		int x2 = anchor.getDx2();
		int y1 = anchor.getDy1();
		int y2 = anchor.getDy2();

		XSSFWorkbook wb = sheet.getWorkbook();
		XSSFCreationHelper newHelper = wb.getCreationHelper();
		XSSFClientAnchor newAnchor = newHelper.createClientAnchor();

		// Row / Column placement.
		if (col1 == -1) {
			newAnchor.setCol1(anchor.getCol1());
		} else {
			newAnchor.setCol1(col1);
		}
		if (col2 == -1) {
			newAnchor.setCol2(anchor.getCol2());
		} else {
			newAnchor.setCol2(col2);
		}
		newAnchor.setRow1(row1);
		newAnchor.setRow2(row2);

		// Fine touch adjustment along the XY coordinate.
		newAnchor.setDx1(x1);
		newAnchor.setDx2(x2);
		newAnchor.setDy1(y1);
		newAnchor.setDy2(y2);

		int newPictureIndex = wb.addPicture(xssfPictureData.getData(), xssfPictureData.getPictureType());

		XSSFPicture newPicture = pat.createPicture(newAnchor, newPictureIndex);
		newPicture.resize(1.0, 1.0);
	}

	protected void copiarContenidoRowByCell(Row rowOrigen, Row rowDestino) {
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

	protected void copiarFormatoRowByCell(Row rowOrigen, Row rowDestino, boolean merged, Sheet sheet, int mergedStars,
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

	protected Cell crearCellConStyle(Row row, int cellNum, CellStyle cellStyle) {
		Cell cell = row.createCell(cellNum);
		cell.setCellStyle(cellStyle);
		return cell;
	}

}
