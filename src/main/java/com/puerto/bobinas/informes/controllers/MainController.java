package com.puerto.bobinas.informes.controllers;

import java.io.File;
import java.io.IOException;
import java.text.DecimalFormat;
import java.util.stream.Collectors;

import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationContext;
import org.springframework.stereotype.Component;

import com.puerto.bobinas.informes.beans.Bobina;
import com.puerto.bobinas.informes.helpers.ExcelHelper;
import com.puerto.bobinas.informes.helpers.TableViewHelper;
import com.puerto.bobinas.informes.tasks.TaskService;

import javafx.fxml.FXML;
import javafx.scene.Cursor;
import javafx.scene.control.Button;
import javafx.scene.control.Hyperlink;
import javafx.scene.control.MenuItem;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TableView;
import javafx.scene.control.TextField;
import javafx.scene.layout.BorderPane;
import javafx.scene.text.Text;
import javafx.scene.text.TextFlow;
import javafx.stage.Stage;
import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class MainController {

	private final ApplicationContext applicationContext;

	@Autowired
	private TableViewHelper tableViewHelper;
	@Autowired
	private ExcelHelper excelHelper;

	@FXML
	private Hyperlink hlResultadoPlantilla;
	@FXML
	private TextFlow textFlowResultadoPlantilla;
	@FXML
	private MenuItem menuItemImportar;
	@FXML
	private Button btnImportar;
	@FXML
	private Button btnPlantilla;
	@FXML
	private ProgressBar pbCargas;
	@FXML
	private TableView<Bobina> tvBobinas;
	@FXML
	private TextField tfRuta;
	@FXML
	private BorderPane bpPrimary;
	@FXML
	private Text txtCliente;
	@FXML
	private Text txtTotalDestinatario;
	@FXML
	private Text txtTotalBobinas;
	@FXML
	private Text txtTotalPesoBruto;

	public MainController(ApplicationContext applicationContext) {
		this.applicationContext = applicationContext;
	}

	@FXML
	public void initialize() {
		textFlowResultadoPlantilla.setVisible(false);
		pbCargas.setVisible(false);
		tfRuta.setEditable(false);
		tvBobinas.setEditable(false);
		btnPlantilla.setDisable(true);
		tableViewHelper.iniciarTableBobinas(tvBobinas);
		//
		limpiarTextosInformativos();
		//
		btnImportar.setOnAction(event -> {
			excelImportar();
		});
		btnPlantilla.setOnAction(event -> {
			excelGenerarSalida();
		});
	}

	@FXML
	public void excelImportar() {
		limpiarTextosInformativos();
		textFlowResultadoPlantilla.setVisible(false);
		hlResultadoPlantilla.setText(StringUtils.EMPTY);
		btnPlantilla.setDisable(true);
		tfRuta.setText(StringUtils.EMPTY);
		tvBobinas.getItems().clear();
		textFlowResultadoPlantilla.setVisible(false);
		Stage primaryStage = (Stage) bpPrimary.getScene().getWindow();
		File selectedFile = excelHelper.getFileChooserEntrada().showOpenDialog(primaryStage);
		if (selectedFile != null) {
			tfRuta.setText(selectedFile.getPath());
			try {
				var servicio = new TaskService();
				servicio.setOnScheduled(event -> {
					initProgressBar();
				});
				servicio.setOnSucceeded(event -> {
					finishProgressBar();
					btnPlantilla.setDisable(false);

				});
				servicio.start();
				var bobinasTemplate = excelHelper.leerExcel(selectedFile.getPath().toString());
				var bobinas = bobinasTemplate.getBobinasList();
				if (!bobinas.isEmpty()) {
					tableViewHelper.rellenarTableBobinas(tvBobinas, bobinas);
					txtCliente.setText(bobinasTemplate.getCliente());
					txtTotalDestinatario.setText(bobinasTemplate.getTotalDestinatarios().toString());
					txtTotalBobinas.setText(bobinasTemplate.getTotalBobinas().toString());
					txtTotalPesoBruto.setText(new DecimalFormat("#,###.#").format(bobinasTemplate.getTotalPeso()));
					;
				} else {
					btnPlantilla.setDisable(true);
				}
			} catch (Exception e) {
				log.error("Error importando Excel", e);
			}
		}
	}

	@FXML
	public void excelGenerarSalida() {
		hlResultadoPlantilla.setText(StringUtils.EMPTY);
		textFlowResultadoPlantilla.setVisible(false);
		var service = new TaskService();
		service = new TaskService();
		service.setOnScheduled(e -> {
			initProgressBar();
			btnPlantilla.setDisable(true);
		});
		service.setOnSucceeded(e -> {
			finishProgressBar();
			btnPlantilla.setDisable(false);
			textFlowResultadoPlantilla.setVisible(true);
		});
		service.restart();
		var bobinas = tvBobinas.getItems().stream().collect(Collectors.toList());
		var rutaSalida = excelHelper.obtenerPlantillaSalida(bobinas);
		hlResultadoPlantilla.setText(rutaSalida.toString());
		hlResultadoPlantilla.setOnAction(event -> {
			btnPlantilla.setDisable(true);
			try {
				Runtime.getRuntime().exec("explorer /select, " + rutaSalida);
			} catch (IOException e) {
				log.error("Error intentando abrir plantilla generada", e);
			}
		});
	}

	@FXML
	public void excelNuevo() {
		limpiarTextosInformativos();
		tfRuta.setText(StringUtils.EMPTY);
		textFlowResultadoPlantilla.setVisible(false);
		pbCargas.setVisible(false);
		tfRuta.setEditable(false);
		tvBobinas.setEditable(false);
		btnPlantilla.setDisable(true);
		tvBobinas.getItems().clear();
	}

	private void finishProgressBar() {
		pbCargas.setVisible(false);
		bpPrimary.setCursor(Cursor.DEFAULT);
	}

	private void initProgressBar() {
		pbCargas.setVisible(true);
		bpPrimary.setCursor(Cursor.WAIT);
	}

	private void limpiarTextosInformativos() {
		txtCliente.setText(StringUtils.EMPTY);
		txtTotalDestinatario.setText(StringUtils.EMPTY);
		txtTotalBobinas.setText(StringUtils.EMPTY);
		txtTotalPesoBruto.setText(StringUtils.EMPTY);
	}
}