package com.puerto.bobinas.informes.stages;

import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.ApplicationListener;
import org.springframework.stereotype.Component;

import com.puerto.bobinas.informes.events.StageExceptionEvent;

import javafx.geometry.HPos;
import javafx.geometry.Insets;
import javafx.geometry.VPos;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Priority;
import javafx.scene.layout.VBox;
import javafx.scene.paint.Color;
import javafx.scene.text.Text;
import javafx.scene.text.TextFlow;
import javafx.stage.Stage;

@Component
public class ExceptionStageInitializer implements ApplicationListener<StageExceptionEvent> {

	@Value("${spring.application.ui.title}")
	private String title;

	private final Label lblTitulo = new Label();
	private final TextFlow tfDescripcion = new TextFlow();

	@Override
	public void onApplicationEvent(StageExceptionEvent event) {
		Stage stage = event.getStage();
		Exception exception = event.getException();
		VBox vbox = new VBox();
		GridPane gp = new GridPane();
		gp.setPadding(new Insets(10));
		gp.setHgap(4);
		gp.setVgap(10);
		VBox.setVgrow(gp, Priority.ALWAYS);
		//
		lblTitulo.setText("Error inesperado aplicacion");
		lblTitulo.setTextFill(Color.RED);
		//
		StringBuilder sb = new StringBuilder();
		sb.append(exception.getClass());
		sb.append(StringUtils.LF);
		sb.append(exception.getMessage());
		tfDescripcion.getChildren().add(new Text(sb.toString()));
		gp.add(lblTitulo, 0, 0);
		gp.add(tfDescripcion, 0, 1);
		GridPane.setConstraints(tfDescripcion, 0, 1, 2, 1, HPos.LEFT, VPos.TOP, Priority.SOMETIMES, Priority.ALWAYS);
		ColumnConstraints col1 = new ColumnConstraints();
		col1.setPercentWidth(100);
		gp.getColumnConstraints().addAll(col1);
		vbox.getChildren().addAll(gp);
		Scene scene = new Scene(vbox);
		stage.setResizable(false);
		stage.setHeight(480);
		stage.setWidth(640);
		stage.setTitle(title);
		stage.setScene(scene);
		stage.show();

	}

}
