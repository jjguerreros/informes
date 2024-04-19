package com.puerto.bobinas.informes.stages;

import java.net.URL;
import java.nio.file.Path;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationListener;
import org.springframework.core.io.Resource;
import org.springframework.stereotype.Component;

import com.puerto.bobinas.informes.constantes.Constantes;
import com.puerto.bobinas.informes.events.StageExceptionEvent;
import com.puerto.bobinas.informes.events.StageReadyEvent;
import com.puerto.bobinas.informes.utils.Utilidades;

import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.scene.image.Image;
import javafx.stage.Stage;
import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class PrimaryStageInitializer implements ApplicationListener<StageReadyEvent> {

	@Value("${spring.application.ui.title}")
	private String title;
	@Value("classpath:views/main.fxml")
	private Resource fxml;
	@Value("classpath:img/icon/icons8-puerto-48.png")
	private Resource appIcon;
	@Value("${fileChooser.directory.root}")
	private String directoryRootPath;
	@Value("${fileChooser.directory.excels.entrada}")
	private String entradaDirectory;
	@Value("${fileChooser.directory.excels.salida}")
	private String salidaDirectory;
	@Autowired
	private Utilidades utilidades;
	private final ApplicationContext applicationContext;

	public PrimaryStageInitializer(ApplicationContext applicationContext) {
		this.applicationContext = applicationContext;
	}

	@Override
	public void onApplicationEvent(StageReadyEvent event) {
		Stage stage = event.getStage();
		try {
			if (utilidades.crearDirectorio(directoryRootPath)) {
				log.info("Directorio {} creado", Path.of(directoryRootPath));
			}
			if (utilidades.crearDirectorio(entradaDirectory)) {
				log.info("Directorio {} creado", Path.of(entradaDirectory));
			}
			if (utilidades.crearDirectorio(salidaDirectory)) {
				log.info("Directorio {} creado", Path.of(salidaDirectory));
			}
			URL url = fxml.getURL();
			FXMLLoader loader = new FXMLLoader(url);
			loader.setControllerFactory(applicationContext::getBean);
			Parent root = loader.load();
			Scene scene = new Scene(root);
			stage.setScene(scene);
			stage.setTitle(title);
			stage.setMinWidth(Constantes.PRIMARY_STAGE_MIN_WIDTH);
			stage.setMinHeight(Constantes.PRIMARY_STAGE_MIN_HEIGHT);
			stage.getIcons().add(new Image(appIcon.getInputStream()));
			stage.show();
		} catch (Exception e) {
			log.error("Error cargando aplicacion", e);
			applicationContext.publishEvent(new StageExceptionEvent(stage, e));
		}

	}

}
