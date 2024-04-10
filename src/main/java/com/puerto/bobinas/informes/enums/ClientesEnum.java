package com.puerto.bobinas.informes.enums;

public enum ClientesEnum {
	THYSSEN("THYSSEN"), ARCELOR("BREMEN");

	private String valor;

	private ClientesEnum(String valor) {
		this.valor = valor;
	}

	public String getValor() {
		return this.valor;
	}

	public static ClientesEnum getClienteEnum(String clienteString) {
		for (ClientesEnum clienteEnum : ClientesEnum.values()) {
			if (clienteEnum.getValor().equals(clienteString)) {
				return clienteEnum;
			}
		}
		return null;
	}
}
