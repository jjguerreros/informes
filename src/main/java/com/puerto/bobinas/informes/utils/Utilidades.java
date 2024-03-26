package com.puerto.bobinas.informes.utils;

import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import lombok.extern.slf4j.Slf4j;

@Component
@Slf4j
public class Utilidades {

	@Value("${user.home}")
	private String userHomePath;

	public boolean crearDirectorio(String pathString) {
		var result = false;
		try {
			var pathStringDirectories = StringUtils.substringAfter(pathString, userHomePath);
			var pathStringDirectoriesArray = StringUtils.split(pathStringDirectories, "/");
			var path = userHomePath + "/";
			for (var pathStringDirectory : pathStringDirectoriesArray) {
				path += pathStringDirectory;
				Path dirPath = Paths.get(path);
				// Si no existe lo creamos
				if (Files.notExists(dirPath)) {
					// Directory not exists
					File directory = new File(path);
					result = directory.mkdir();
				}
				path += "/";
			}
		} catch (Exception e) {
			log.error("Error creando directorios", e);
			return false;
		}
		return result;
	}

	public String obtenerFechaString(Date date, String patternDate) {
		try {
			DateFormat df = new SimpleDateFormat(patternDate);
			return df.format(date);
		} catch (Exception e) {
			return StringUtils.EMPTY;
		}
	}

}
