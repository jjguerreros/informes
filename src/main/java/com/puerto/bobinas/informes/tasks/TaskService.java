package com.puerto.bobinas.informes.tasks;

import javafx.concurrent.Service;
import javafx.concurrent.Task;

public class TaskService extends Service<Void> {

	@Override
	protected Task<Void> createTask() {
		Task<Void> task = new Task<Void>() {

			@Override
			protected Void call() throws Exception {
				Thread.sleep(3000);

//				for (int p = 0; p < 100; p++) {
//					updateProgress(p, 100);
//				}
				return null;
			}
		};
		return task;
	}
}
