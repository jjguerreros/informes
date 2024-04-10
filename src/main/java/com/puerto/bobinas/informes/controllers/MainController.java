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
import com.puerto.bobinas.informes.beans.BobinasTemplate;
import com.puerto.bobinas.informes.enums.ClientesEnum;
import com.puerto.bobinas.informes.helpers.ExcelHelper;
import com.puerto.bobinas.informes.helpers.TableViewHelper;
import com.puerto.bobinas.informes.tasks.TaskService;

import javafx.fxml.FXML;
import javafx.scene.Cursor;
import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.control.Button;
import javafx.scene.control.ChoiceBox;
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
	private TextField tfBArco;
	@FXML
	private BorderPane bpPrimary;
	@FXML
	private Text txtEncabezado;
	@FXML
	private ChoiceBox<ClientesEnum> cbCliente;
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
		cbCliente.getItems().addAll(ClientesEnum.values());
		//
		limpiarTextosInformativos();
		//
		btnImportar.setOnAction(event -> {
			excelImportar();
		});
		btnPlantilla.setOnAction(event -> {
			excelGenerarSalida();
		});
		//
	}

	@FXML
	public void excelImportar() {
		limpiarTextosInformativos();
		textFlowResultadoPlantilla.setVisible(false);
		hlResultadoPlantilla.setText(StringUtils.EMPTY);
		btnPlantilla.setDisable(true);
		tvBobinas.getItems().clear();
		textFlowResultadoPlantilla.setVisible(false);
		Stage primaryStage = (Stage) bpPrimary.getScene().getWindow();
		File selectedFile = excelHelper.getFileChooserEntrada().showOpenDialog(primaryStage);
		if (selectedFile != null) {
			tfRuta.setText(selectedFile.getPath());
			try {
				var servicio = new TaskService();
				var bobinasTemplate = excelHelper.getBobinasTemplate(selectedFile.getPath().toString());
				cbCliente.setValue(ClientesEnum.getClienteEnum(bobinasTemplate.getCliente()));
				tfBArco.setText(bobinasTemplate.getBarco());
				//
				txtTotalDestinatario.setText(bobinasTemplate.getTotalDestinatarios().toString());
				txtTotalBobinas.setText(bobinasTemplate.getTotalBobinas().toString());
				txtTotalPesoBruto.setText(new DecimalFormat("#,###.#").format(bobinasTemplate.getTotalPeso()));
				txtEncabezado.setText(bobinasTemplate.getEncabezado());
				//
				var bobinas = bobinasTemplate.getBobinasList();
				if (!bobinas.isEmpty()) {
					servicio.setOnScheduled(event -> {
						initProgressBar();
						cbCliente.setDisable(true);
						tfBArco.setDisable(true);
						tableViewHelper.rellenarTableBobinas(tvBobinas, bobinas);
					});
					servicio.setOnSucceeded(event -> {
						finishProgressBar();
						cbCliente.setDisable(false);
						tfBArco.setDisable(false);
						btnPlantilla.setDisable(false);

					});
					servicio.start();
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
		String barcoVal = tfBArco.getText();
		// TODO: Validaciones
		if (StringUtils.isBlank(barcoVal)) {
			Alert alert = new Alert(AlertType.WARNING);
			alert.setTitle("Errores formulario");
			alert.setContentText("\"Barco\" no puede estar vacio");
			alert.setHeaderText(null);
			alert.initOwner(bpPrimary.getScene().getWindow());
			alert.show();
			return;
		}
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
		var bobinasTemplate = new BobinasTemplate();
		var bobinas = tvBobinas.getItems().stream().collect(Collectors.toList());
		bobinasTemplate.setBobinasList(bobinas);
		bobinasTemplate.setCliente(cbCliente.getValue().getValor());
		bobinasTemplate.setBarco(barcoVal);
		var rutaSalida = excelHelper.obtenerPlantillaSalida(bobinasTemplate);
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
		txtEncabezado.setText(StringUtils.EMPTY);
		cbCliente.setValue(null);
		txtTotalDestinatario.setText(StringUtils.EMPTY);
		txtTotalBobinas.setText(StringUtils.EMPTY);
		txtTotalPesoBruto.setText(StringUtils.EMPTY);
		tfBArco.setText(StringUtils.EMPTY);
		tfRuta.setText(StringUtils.EMPTY);
	}
}