import locale
import logging
import sys
from datetime import datetime

import pandas as pd

from estados import ESTADOS

logging.basicConfig(
    level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s"
)


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
        locale.setlocale(locale.LC_ALL, 'pt_BR')
        if pd.isna(data):
            return data
        try:
            if isinstance(data, (pd.Timestamp, datetime)):
                data_obj = pd.to_datetime(data)
                return data_obj.strftime("%d/%b/%Y").capitalize()
        except ValueError:
            return None

    @staticmethod
    def formatar_valor(valor):
        if pd.isna(valor):
            return valor
        try:
            valor_str = str(valor)
            if "," in valor_str and len(valor_str.split(",")[-1]) == 2:
                return valor
            valor_limpo = (
                str(valor)
                .replace(",", "")
            )

            valor_inteiro = int(valor_limpo)

            valor_float = valor_inteiro / 100

            return (
                f"{valor_float:,.2f}"
                .replace(".", ",")
            )
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

            df_final.to_csv(self.output_file, index=False, sep=";", encoding='utf-8')
            logging.info(f"Relatório gerado com sucesso no arquivo {
                         self.output_file}.")

        except Exception as e:
            logging.error(f"Erro ao processar o arquivo: {e}")
            sys.exit(1)
