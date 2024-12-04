# Anais dos Eventos Científicos

Este projeto automatiza a criação dos anais de eventos científicos a partir de uma planilha Excel contendo os resumos, autores, e informações relevantes. Os anais são gerados no formato Word com cabeçalho personalizado, resumo justificado e formatação apropriada.

## ⚙️ Requisitos

    Python 3.7+
    Bibliotecas:
        pandas
        python-docx
        beautifulsoup4

Instale as dependências com:

    pip install -r requirements.txt

## 🚀 Como Usar

Adicione o arquivo Excel e o logotipo no diretório data/.
Execute o script:

    python src/anais.py --input data/2023_Anais.xlsx --output outputs/anais_formatado.docx

O documento gerado será salvo na pasta outputs/.


## 📁 Estrutura do Projeto

```plaintext
/anaisdoseventos
├── data/                # Arquivos de entrada (Excel e imagem)
├── src/                 # Código fonte do projeto
├── outputs/             # Anais gerados
├── README.md            # Documentação do projeto
└── requirements.txt     # Dependências do projeto
