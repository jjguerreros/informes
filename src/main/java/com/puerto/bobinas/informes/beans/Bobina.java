package com.puerto.bobinas.informes.beans;

import java.util.Objects;

import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;

import javafx.beans.property.DoubleProperty;
import javafx.beans.property.SimpleDoubleProperty;
import javafx.beans.property.SimpleStringProperty;
import javafx.beans.property.StringProperty;

public class Bobina {

	private StringProperty numSerie;
	private StringProperty nombreDestinatario;
	private DoubleProperty pesoBrutoPrevisto;

	public Bobina() {
		super();
		numSerie = new SimpleStringProperty(StringUtils.EMPTY);
		nombreDestinatario = new SimpleStringProperty(StringUtils.EMPTY);
		pesoBrutoPrevisto = new SimpleDoubleProperty(NumberUtils.DOUBLE_ZERO);
	}

	public Bobina(String numSerie, String nombreDestinatario, double pesoBrutoPrevisto) {
		super();
		this.numSerie = new SimpleStringProperty(numSerie);
		this.nombreDestinatario = new SimpleStringProperty(nombreDestinatario);
		this.pesoBrutoPrevisto = new SimpleDoubleProperty(pesoBrutoPrevisto);
	}

	public StringProperty getNumSerieProperty() {
		return numSerie;
	}

	public String getNumSerie() {
		return numSerie.get();
	}

	public void setNumSerie(String numSerie) {
		this.numSerie.set(numSerie);
	}

	public StringProperty getNombreDestinatarioProperty() {
		return nombreDestinatario;
	}

	public String getNombreDestinatario() {
		return nombreDestinatario.get();
	}

	public void setNombreDestinatario(String nombreDestinatario) {
		this.nombreDestinatario.set(nombreDestinatario);
	}

	public DoubleProperty getPesoBrutoPrevistoProperty() {
		return pesoBrutoPrevisto;
	}

	public Double getPesoBrutoPrevisto() {
		return pesoBrutoPrevisto.get();
	}

	public void setPesoBrutoPrevisto(double pesoBrutoPrevisto) {
		this.pesoBrutoPrevisto.set(pesoBrutoPrevisto);
	}

	@Override
	public boolean equals(final Object obj) {
		if (this == obj) {
			return true;
		}
		if (obj == null) {
			return false;
		}
		if (getClass() != obj.getClass()) {
			return false;
		}
		Bobina other = (Bobina) obj;
		return Objects.equals(nombreDestinatario.getValue(), other.nombreDestinatario.getValue());
	}

}
