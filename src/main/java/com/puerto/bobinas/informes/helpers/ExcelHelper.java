package com.puerto.bobinas.informes.helpers;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.nio.file.LinkOption;
import java.nio.file.Path;
import java.text.DecimalFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.stream.Collectors;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import com.puerto.bobinas.informes.beans.ArcelorPlantilla;
import com.puerto.bobinas.informes.beans.Bobina;
import com.puerto.bobinas.informes.beans.BobinasTemplate;
import com.puerto.bobinas.informes.beans.ThyssenPlantilla;
import com.puerto.bobinas.informes.constantes.Constantes;
import com.puerto.bobinas.informes.enums.ClientesEnum;

import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class ExcelHelper {

	@Autowired
	private ThyssenPlantilla thyssenPlantilla;
	@Autowired
	private ArcelorPlantilla arcelorPlantilla;

	@Value("${fileChooser.directory.root}")
	private String rootDirectory;
	@Value("${fileChooser.directory.excels.entrada}")
	private String entradaDirectory;
	@Value("${fileChooser.dialog.title}")
	private String fileChooserDialogTitle;
	@Value("${tableView.bobinas.serie.alias}")
	private String serieAlias;
	@Value("${tableView.bobinas.destinatario.alias}")
	private String destinatarioAlias;
	@Value("${tableView.bobinas.pesoBruto.alias}")
	private String pesoBrutoAlias;

	public FileChooser getFileChooserEntrada() {
		final FileChooser fileChooser = new FileChooser();
		ExtensionFilter ex1 = new ExtensionFilter("Excel File", "*.xlsx", "*.xls");
		ExtensionFilter ex2 = new ExtensionFilter("all Files", "*.*");
		fileChooser.getExtensionFilters().addAll(ex1, ex2);
		fileChooser.setTitle(fileChooserDialogTitle);
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
			var clienteEnum = ClientesEnum.getClienteEnum(bobinasTemplate.getCliente());
			switch (clienteEnum) {
			case THYSSEN:
				return thyssenPlantilla.generarPlantilla(bobinasTemplate);
			default:
				return arcelorPlantilla.generarPlantilla(bobinasTemplate);
			}
		} catch (Exception ex) {
			log.error("Error generando plantilla", ex);
			return Path.of(rootDirectory);
		}
	}
}
