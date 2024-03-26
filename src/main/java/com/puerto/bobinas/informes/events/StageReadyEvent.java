package com.puerto.bobinas.informes.events;

import org.springframework.context.ApplicationEvent;

import javafx.stage.Stage;

public class StageReadyEvent extends ApplicationEvent {
	/**
	 * 
	 */
	private static final long serialVersionUID = -7041896357848714557L;

	public StageReadyEvent(Stage stage) {
		super(stage);
	}

	public Stage getStage() {
		return Stage.class.cast(getSource());
	}

}
