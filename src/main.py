import logging
import sys

from relatorio import RelatorioGenerator


def main():

    if len(sys.argv) != 3:
        logging.error(
            "Uso incorreto! Exemplo: python main.py arquivo_entrada.xlsx relatorio_gerado.csv"
        )
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]

    try:
        relatorio = RelatorioGenerator(input_file, output_file)
        relatorio.processar_dados()
    except Exception as e:
        logging.error(f"Falha ao gerar o relatório: {e}")


if __name__ == "__main__":
    main()

# execução: python src/main.py data/arquivo_entrada.xlsx data/relatorio_gerado.csv
