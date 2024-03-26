package com.puerto.bobinas.informes.events;

import org.springframework.context.ApplicationEvent;

import javafx.stage.Stage;

public class StageExceptionEvent extends ApplicationEvent {

	/**
	 * 
	 */
	private static final long serialVersionUID = -3659716888723307095L;
	private final Exception exception;

	public StageExceptionEvent(Stage stage, Exception exception) {
		super(stage);
		this.exception = exception;
	}

	public Stage getStage() {
		return Stage.class.cast(getSource());
	}

	public Exception getException() {
		return exception;
	}

}
