# 📊 Projeto de Análise de Fluxos com Excel

    Este projeto tem como objetivo automatizar a leitura de uma planilha Excel com dados de fluxos, gerar uma nova planilha com os resultados mensais e criar gráficos consolidados com base nesses dados.

---
## 📁 Estrutura do Projeto

    fluxo_analise/ │ 
        ├── main.py # Script principal 
        ├── config/ │ 
            └── settings.py # Caminhos e configurações fixas 
        ├── modules/ │ 
            ├── gerar_planilha.py # Função: gerar_fluxo_mensal(mes) 
            │ └── gerar_graficos.py # Função: gerar_graficos_gerais() 
        ├── utils/ │ 
            └── excel_helpers.py # Funções auxiliares 
        ├── data/ # Planilhas de entrada/saída (opcional) 
        ├── .gitignore # Arquivos ignorados pelo Git 
        ├── requirements.txt # Dependências do projeto 
        └── README.md # Este arquivo
---

## ▶️ Como Executar

1. **Crie e ative o ambiente virtual:**

    ```bash
    python -m venv .venv
    source .venv/bin/activate  # Linux/macOS
    .venv\Scripts\activate     # Windows 
    
2. Instale as dependências:
    pip install -r requirements.txt

3. Execute o projeto:
    python main.py


⚙️ Configurações
    Os caminhos das planilhas e o nome da aba modelo estão definidos no arquivo:

    config/settings.py
    
🧪 Testes
    (Testes automatizados podem ser adicionados futuramente na pasta tests/.)

📌 Requisitos
    Python 3.9+
    pandas
    openpyxl
    
📬 Contribuições
    Sinta-se à vontade para sugerir melhorias ou abrir issues!

