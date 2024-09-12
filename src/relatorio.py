import logging
import sys
from datetime import datetime

import pandas as pd

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)

ESTADOS = {
    "AC": "Acre",
    "AL": "Alagoas",
    "AP": "Amapá",
    "AM": "Amazonas",
    "BA": "Bahia",
    "CE": "Ceará",
    "DF": "Distrito Federal",
    "ES": "Espírito Santo",
    "GO": "Goiás",
    "MA": "Maranhão",
    "MT": "Mato Grosso",
    "MS": "Mato Grosso do Sul",
    "MG": "Minas Gerais",
    "PA": "Pará",
    "PB": "Paraíba",
    "PR": "Paraná",
    "PE": "Pernambuco",
    "PI": "Piauí",
    "RJ": "Rio de Janeiro",
    "RN": "Rio Grande do Norte",
    "RS": "Rio Grande do Sul",
    "RO": "Rondônia",
    "RR": "Roraima",
    "SC": "Santa Catarina",
    "SP": "São Paulo",
    "SE": "Sergipe",
    "TO": "Tocantins",
}


class RelatorioGenerator:
    def __init__(self, input_file, output_file):
        self.input_file = input_file
        self.output_file = output_file

    @staticmethod
    def remover_formatacao(valor):
        if pd.isna(valor):
            return valor
        return str(valor).replace(".", "").replace("/", "").replace("-", "")

    @staticmethod
    def traduzir_estado(sigla):
        return ESTADOS.get(sigla, sigla)

    @staticmethod
    def formatar_data(data):
        if pd.isna(data):
            return data
        try:
            # Se data é um Timestamp ou datetime, converte diretamente para o formato desejado
            if isinstance(data, (pd.Timestamp, datetime)):
                data_obj = pd.to_datetime(data)  # Garante que é um objeto datetime
            else:
                # Converte a string para um objeto datetime
                data_obj = pd.to_datetime(data, format="%d/%m/%Y", errors="coerce")
            # Formata a data para o formato desejado
            return (
                data_obj.strftime("%d/%B/%Y").capitalize()
                if not pd.isna(data_obj)
                else None
            )
        except ValueError:
            return None

    @staticmethod
    def formatar_valor(valor):
        if pd.isna(valor):
            return valor
        try:
            return f"{float(valor):,.2f}".replace(".", ",")
        except ValueError:
            return valor

    def processar_dados(self):
        logging.info(f"Lendo o arquivo {self.input_file}...")
        try:
            df = pd.read_excel(self.input_file)
            logging.info("Arquivo lido com sucesso.")

            colunas_originais = df.columns.tolist()

            processadores = {
                1: self.remover_formatacao,  # Coluna 2
                6: self.traduzir_estado,  # Coluna 7
                9: self.formatar_data,  # Coluna 10
                10: self.formatar_valor,  # Coluna 11
            }

            for col_idx, func in processadores.items():
                if col_idx < len(colunas_originais):
                    df[colunas_originais[col_idx]] = df[
                        colunas_originais[col_idx]
                    ].apply(func)

            df_final = df[colunas_originais]

            df_final = df_final.sort_values(by=colunas_originais[7])

            df_final.to_csv(self.output_file, index=False, sep=";")
            logging.info(f"Relatório gerado com sucesso no arquivo {self.output_file}.")

        except Exception as e:
            logging.error(f"Erro ao processar o arquivo: {e}")
            sys.exit(1)
