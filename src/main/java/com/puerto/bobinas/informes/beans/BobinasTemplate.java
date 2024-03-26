package com.puerto.bobinas.informes.beans;

import java.util.ArrayList;
import java.util.List;

public class BobinasTemplate {
	private String cliente;
	private Integer totalBobinas;
	private Integer totalDestinatarios;
	private Double totalPeso;
	private List<Bobina> bobinasList;

	public String getCliente() {
		return cliente;
	}

	public void setCliente(String cliente) {
		this.cliente = cliente;
	}

	public Integer getTotalBobinas() {
		return totalBobinas;
	}

	public void setTotalBobinas(Integer totalBobinas) {
		this.totalBobinas = totalBobinas;
	}

	public Integer getTotalDestinatarios() {
		return totalDestinatarios;
	}

	public void setTotalDestinatarios(Integer totalDestinatarios) {
		this.totalDestinatarios = totalDestinatarios;
	}

	public Double getTotalPeso() {
		return totalPeso;
	}

	public void setTotalPeso(Double totalPeso) {
		this.totalPeso = totalPeso;
	}

	public List<Bobina> getBobinasList() {
		if (null == bobinasList) {
			bobinasList = new ArrayList<Bobina>();
		}
		return bobinasList;
	}

	public void setBobinasList(List<Bobina> bobinasList) {
		this.bobinasList = bobinasList;
	}

}
