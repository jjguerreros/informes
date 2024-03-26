package com.puerto.bobinas.informes.helpers;

import java.util.List;

import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import com.puerto.bobinas.informes.beans.Bobina;

import javafx.event.EventHandler;
import javafx.scene.control.TableColumn;
import javafx.scene.control.TableColumn.CellEditEvent;
import javafx.scene.control.TableView;
import javafx.scene.control.cell.PropertyValueFactory;
import javafx.scene.control.cell.TextFieldTableCell;
import javafx.util.converter.DoubleStringConverter;

@Component
public class TableViewHelper {

	@Value("${tableView.bobinas.serie}")
	private String numSerie;
	@Value("${tableView.bobinas.destinatario}")
	private String destinario;
	@Value("${tableView.bobinas.pesoBruto}")
	private String pesoBruto;

	public void iniciarTableBobinas(TableView<Bobina> tableView) {
		TableColumn<Bobina, String> tableColumnSerie = new TableColumn<Bobina, String>(numSerie);
		tableColumnSerie.setCellValueFactory(new PropertyValueFactory<Bobina, String>("numSerie"));
		tableColumnSerie.setCellFactory(TextFieldTableCell.forTableColumn());
		tableColumnSerie.setOnEditCommit(new EventHandler<CellEditEvent<Bobina, String>>() {
			@Override
			public void handle(CellEditEvent<Bobina, String> event) {
				Bobina bobina = event.getRowValue();
				bobina.setNumSerie(event.getNewValue());
			}
		});
		TableColumn<Bobina, String> tableColumnDestinatario = new TableColumn<Bobina, String>(destinario);
		tableColumnDestinatario.setCellValueFactory(new PropertyValueFactory<Bobina, String>("nombreDestinatario"));
		tableColumnDestinatario.setCellFactory(TextFieldTableCell.forTableColumn());
		tableColumnDestinatario.setOnEditCommit(new EventHandler<CellEditEvent<Bobina, String>>() {
			@Override
			public void handle(CellEditEvent<Bobina, String> event) {
				Bobina bobina = event.getRowValue();
				bobina.setNombreDestinatario(event.getNewValue());
			}
		});
		TableColumn<Bobina, Double> tableColumnPesoBruto = new TableColumn<Bobina, Double>(pesoBruto);
		tableColumnPesoBruto.setCellValueFactory(new PropertyValueFactory<Bobina, Double>("pesoBrutoPrevisto"));
		tableColumnPesoBruto.setCellFactory(TextFieldTableCell.forTableColumn(new DoubleStringConverter()));
		tableColumnPesoBruto.setOnEditCommit(new EventHandler<CellEditEvent<Bobina, Double>>() {
			@Override
			public void handle(CellEditEvent<Bobina, Double> event) {
				Bobina bobina = event.getRowValue();
				bobina.setPesoBrutoPrevisto(event.getNewValue());
			}
		});
		tableView.getColumns().add(tableColumnDestinatario);
		tableView.getColumns().add(tableColumnSerie);
		tableView.getColumns().add(tableColumnPesoBruto);
	}

	public void rellenarTableBobinas(TableView<Bobina> tvBobinas, List<Bobina> bobinas) {
		bobinas.forEach(bobina -> tvBobinas.getItems().add(bobina));
		ordenarTableByPosCol(tvBobinas, 0);
	}

	public void ordenarTableByPosCol(TableView<Bobina> tableView, int posCol) {
		if(tableView.getSortOrder().isEmpty()) {
			tableView.getSortOrder().add(tableView.getColumns().get(posCol));
		}
		tableView.sort();
	}

}
